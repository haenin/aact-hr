package com.aact.overtime.service;

import com.aact.overtime.entity.DepartmentSummary;
import com.aact.overtime.entity.OvertimeRecord;
import com.aact.overtime.repository.DepartmentSummaryRepository;
import com.aact.overtime.repository.OvertimeRecordRepository;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Component;

import java.io.ByteArrayOutputStream;
import java.util.*;
import java.util.stream.Collectors;

/**
 * DB 데이터 → 시간외근무수당 신청서 엑셀 생성
 *
 * 시트 구성:
 *  ├── {부서명}표지  (부서별 집계표 + 결재란)
 *  └── {부서명}({월}) (개인별 세부내역)
 *
 * 결재란은 approvers 리스트로 동적 처리
 * ex) ["담당", "과장", "상무", "부사장"]
 */
@Component
@RequiredArgsConstructor
public class ExcelSheetGenerator {

    private final OvertimeRecordRepository overtimeRecordRepository;
    private final DepartmentSummaryRepository departmentSummaryRepository;

    /**
     * 연월 기준으로 전체 엑셀 생성
     *
     * @param applyYearMonth "2026-04"
     * @param approvers      결재란 직급 목록 ex) ["담당", "과장", "상무", "부사장"] (동적)
     * @return 엑셀 파일 byte[]
     */
    public byte[] generate(String applyYearMonth, List<String> approvers) throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook();
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {

            List<DepartmentSummary> summaries = departmentSummaryRepository
                    .findByApplyYearMonthOrderByDepartmentAsc(applyYearMonth);
            List<OvertimeRecord> records = overtimeRecordRepository
                    .findByApplyYearMonthOrderByDepartmentAscEmployeeNameAsc(applyYearMonth);

            Map<String, List<DepartmentSummary>> summaryByDept = summaries.stream()
                    .collect(Collectors.groupingBy(DepartmentSummary::getDepartment));
            Map<String, List<OvertimeRecord>> recordByDept = records.stream()
                    .collect(Collectors.groupingBy(OvertimeRecord::getDepartment));

            List<String> departments = summaries.stream()
                    .map(DepartmentSummary::getDepartment)
                    .distinct()
                    .collect(Collectors.toList());

            StyleSet styles = new StyleSet(wb);

            for (String dept : departments) {
                createCoverSheet(wb, styles, applyYearMonth, dept,
                        summaryByDept.getOrDefault(dept, Collections.emptyList()),
                        approvers != null ? approvers : Collections.emptyList());

                List<OvertimeRecord> deptRecords = recordByDept.getOrDefault(dept, Collections.emptyList());
                if (!deptRecords.isEmpty()) {
                    createDetailSheet(wb, styles, applyYearMonth, dept, deptRecords);
                }
            }

            wb.write(out);
            return out.toByteArray();
        }
    }

    // ─────────────────────────────────────────────────────────
    // 표지 시트
    // ─────────────────────────────────────────────────────────
    private void createCoverSheet(XSSFWorkbook wb, StyleSet styles,
                                  String applyYearMonth, String dept,
                                  List<DepartmentSummary> summaries,
                                  List<String> approvers) {
        String month = applyYearMonth.substring(5);
        XSSFSheet sheet = wb.createSheet(dept + "표지");

        // ── row 0: 제목
        Row r0 = sheet.createRow(0);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 34));
        Cell title = r0.createCell(0);
        title.setCellValue("시간 외 근무수당 신청서(" + month + "월)");
        title.setCellStyle(styles.title);

        // ── 결재란 (동적 - List<String> approvers)
        if (!approvers.isEmpty()) {
            Row approvalHeaderRow = sheet.createRow(3);
            setCell(approvalHeaderRow, 16, "결재", styles.header);

            int startCol = 17;
            int colSpan  = 3;

            for (int i = 0; i < approvers.size(); i++) {
                int col = startCol + (i * colSpan);
                sheet.addMergedRegion(new CellRangeAddress(3, 3, col, col + colSpan - 1));
                setCell(approvalHeaderRow, col, approvers.get(i), styles.header);

                // 서명칸 (row 4~7)
                for (int r = 4; r <= 7; r++) {
                    Row sigRow = getOrCreateRow(sheet, r);
                    sheet.addMergedRegion(new CellRangeAddress(r, r, col, col + colSpan - 1));
                    sigRow.createCell(col).setCellStyle(styles.data);
                }
            }
        }

        // ── row 7: 섹션 제목
        Row r7 = getOrCreateRow(sheet, 7);
        r7.createCell(0).setCellValue("1. 부서별 집계표");

        // ── row 9: 헤더 1행
        Row r9 = sheet.createRow(9);
        sheet.addMergedRegion(new CellRangeAddress(9, 10, 0, 0));
        sheet.addMergedRegion(new CellRangeAddress(9, 9, 1, 3));
        sheet.addMergedRegion(new CellRangeAddress(9, 9, 4, 6));
        sheet.addMergedRegion(new CellRangeAddress(9, 9, 7, 9));
        sheet.addMergedRegion(new CellRangeAddress(9, 10, 33, 34));
        setCell(r9, 0,  "부서명", styles.header);
        setCell(r9, 1,  "전월(" + prevMonth(applyYearMonth) + "월)", styles.header);
        setCell(r9, 4,  "당월(" + month + "월)", styles.header);
        setCell(r9, 7,  "증감", styles.header);
        setCell(r9, 33, "증감사유", styles.header);

        // ── row 10: 헤더 2행
        Row r10 = sheet.createRow(10);
        for (int sc : new int[]{1, 4, 7}) {
            setCell(r10, sc,     "연장", styles.header);
            setCell(r10, sc + 1, "야간", styles.header);
            setCell(r10, sc + 2, "휴일", styles.header);
        }

        // ── 데이터 행
        int rowIdx = 11;
        double totalPrevExt = 0, totalPrevNight = 0, totalPrevHol = 0;
        double totalCurrExt = 0, totalCurrNight = 0, totalCurrHol = 0;

        for (DepartmentSummary s : summaries) {
            Row dr = sheet.createRow(rowIdx++);
            setCell(dr, 0, s.getDepartment(), styles.data);
            setNumCell(dr, 1, s.getPrevExtensionHours(), styles.data);
            setNumCell(dr, 2, s.getPrevNightHours(), styles.data);
            setNumCell(dr, 3, s.getPrevHolidayHours(), styles.data);
            setNumCell(dr, 4, s.getCurrExtensionHours(), styles.data);
            setNumCell(dr, 5, s.getCurrNightHours(), styles.data);
            setNumCell(dr, 6, s.getCurrHolidayHours(), styles.data);
            setNumCell(dr, 7, s.getDiffExtensionHours(), styles.data);
            setNumCell(dr, 8, s.getDiffNightHours(), styles.data);
            setNumCell(dr, 9, s.getDiffHolidayHours(), styles.data);
            setCell(dr, 33, s.getChangeReason() != null ? s.getChangeReason() : "", styles.data);

            totalPrevExt   += nvl(s.getPrevExtensionHours());
            totalPrevNight += nvl(s.getPrevNightHours());
            totalPrevHol   += nvl(s.getPrevHolidayHours());
            totalCurrExt   += nvl(s.getCurrExtensionHours());
            totalCurrNight += nvl(s.getCurrNightHours());
            totalCurrHol   += nvl(s.getCurrHolidayHours());
        }

        // ── 총 합계
        Row totalRow = sheet.createRow(rowIdx++);
        setCell(totalRow, 0, "총 합계", styles.header);
        setNumCell(totalRow, 1, totalPrevExt,   styles.header);
        setNumCell(totalRow, 2, totalPrevNight, styles.header);
        setNumCell(totalRow, 3, totalPrevHol,   styles.header);
        setNumCell(totalRow, 4, totalCurrExt,   styles.header);
        setNumCell(totalRow, 5, totalCurrNight, styles.header);
        setNumCell(totalRow, 6, totalCurrHol,   styles.header);
        setNumCell(totalRow, 7, totalCurrExt   - totalPrevExt,   styles.header);
        setNumCell(totalRow, 8, totalCurrNight - totalPrevNight, styles.header);
        setNumCell(totalRow, 9, totalCurrHol   - totalPrevHol,   styles.header);

        // ── 안내 문구
        rowIdx += 2;
        Row noteRow = sheet.createRow(rowIdx++);
        sheet.addMergedRegion(new CellRangeAddress(rowIdx - 1, rowIdx - 1, 0, 34));
        noteRow.createCell(0).setCellValue(
                "\u3000\u3000상기와 같이 " + month + "월 시간외 수당을 신청하오니 검토하시고 재가하여 주시기 바랍니다.");

        // ── 날짜
        rowIdx += 10;
        Row dateRow = sheet.createRow(rowIdx++);
        sheet.addMergedRegion(new CellRangeAddress(rowIdx - 1, rowIdx - 1, 0, 34));
        dateRow.createCell(0).setCellValue(applyYearMonth.replace("-", ". ") + ". 07");

        // ── 회사명
        rowIdx += 2;
        Row compRow = sheet.createRow(rowIdx);
        sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx, 0, 34));
        Cell compCell = compRow.createCell(0);
        compCell.setCellValue("에 이 에 이 씨 티 유 한 회 사");
        compCell.setCellStyle(styles.title);
    }

    // ─────────────────────────────────────────────────────────
    // 세부내역 시트
    // ─────────────────────────────────────────────────────────
    private void createDetailSheet(XSSFWorkbook wb, StyleSet styles,
                                   String applyYearMonth, String dept,
                                   List<OvertimeRecord> records) {
        String month = applyYearMonth.substring(5);
        XSSFSheet sheet = wb.createSheet(dept + "(" + month + "월)");

        Row r0 = sheet.createRow(0);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 12));
        r0.createCell(0).setCellValue(month + "월 시간 외 근로시간 개인별 세부내역 (" + dept + ")");

        Row r2 = sheet.createRow(2);
        sheet.addMergedRegion(new CellRangeAddress(2, 3, 0, 0));
        sheet.addMergedRegion(new CellRangeAddress(2, 3, 1, 1));
        sheet.addMergedRegion(new CellRangeAddress(2, 3, 2, 2));
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 3, 4));
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 5, 6));
        sheet.addMergedRegion(new CellRangeAddress(2, 3, 7, 7));
        sheet.addMergedRegion(new CellRangeAddress(2, 3, 8, 8));
        sheet.addMergedRegion(new CellRangeAddress(2, 3, 9, 9));
        sheet.addMergedRegion(new CellRangeAddress(2, 3, 10, 10));
        setCell(r2, 0,  "부서",          styles.header);
        setCell(r2, 1,  "성명",          styles.header);
        setCell(r2, 2,  "일자",          styles.header);
        setCell(r2, 3,  "예정근무",      styles.header);
        setCell(r2, 5,  "실근무(지문)",  styles.header);
        setCell(r2, 7,  "연장",          styles.header);
        setCell(r2, 8,  "야간",          styles.header);
        setCell(r2, 9,  "휴일",          styles.header);
        setCell(r2, 10, "휴일연장",      styles.header);

        Row r3 = sheet.createRow(3);
        setCell(r3, 3, "출근", styles.header);
        setCell(r3, 4, "퇴근", styles.header);
        setCell(r3, 5, "출근", styles.header);
        setCell(r3, 6, "퇴근", styles.header);

        Map<String, List<OvertimeRecord>> byEmployee = new LinkedHashMap<>();
        for (OvertimeRecord r : records) {
            byEmployee.computeIfAbsent(r.getEmployeeName(), k -> new ArrayList<>()).add(r);
        }

        int rowIdx = 4;
        double grandExt = 0, grandNight = 0, grandHol = 0, grandHolExt = 0;

        for (Map.Entry<String, List<OvertimeRecord>> entry : byEmployee.entrySet()) {
            String empName = entry.getKey();
            List<OvertimeRecord> empRecords = entry.getValue();
            double subExt = 0, subNight = 0, subHol = 0, subHolExt = 0;
            boolean firstRow = true;

            for (OvertimeRecord rec : empRecords) {
                Row dr = sheet.createRow(rowIdx++);
                if (firstRow) {
                    setCell(dr, 0, rec.getDepartment(), styles.data);
                    setCell(dr, 1, empName, styles.data);
                    firstRow = false;
                }
                setCell(dr, 2, rec.getWorkDate()       != null ? rec.getWorkDate().toString()       : "", styles.data);
                setCell(dr, 3, rec.getScheduledStart() != null ? rec.getScheduledStart().toString() : "", styles.data);
                setCell(dr, 4, rec.getScheduledEnd()   != null ? rec.getScheduledEnd().toString()   : "", styles.data);
                setCell(dr, 5, rec.getActualStart()    != null ? rec.getActualStart().toString()    : "", styles.data);
                setCell(dr, 6, rec.getActualEnd()      != null ? rec.getActualEnd().toString()      : "", styles.data);
                setNumCell(dr, 7,  rec.getExtensionHours(),        styles.data);
                setNumCell(dr, 8,  rec.getNightHours(),            styles.data);
                setNumCell(dr, 9,  rec.getHolidayHours(),          styles.data);
                setNumCell(dr, 10, rec.getHolidayExtensionHours(), styles.data);

                subExt    += nvl(rec.getExtensionHours());
                subNight  += nvl(rec.getNightHours());
                subHol    += nvl(rec.getHolidayHours());
                subHolExt += nvl(rec.getHolidayExtensionHours());
            }

            Row subRow = sheet.createRow(rowIdx++);
            setCell(subRow, 1, "소계", styles.header);
            setNumCell(subRow, 7,  subExt,    styles.header);
            setNumCell(subRow, 8,  subNight,  styles.header);
            setNumCell(subRow, 9,  subHol,    styles.header);
            setNumCell(subRow, 10, subHolExt, styles.header);

            grandExt    += subExt;
            grandNight  += subNight;
            grandHol    += subHol;
            grandHolExt += subHolExt;
        }

        Row totalRow = sheet.createRow(rowIdx);
        setCell(totalRow, 1, "합계", styles.header);
        setNumCell(totalRow, 7,  grandExt,    styles.header);
        setNumCell(totalRow, 8,  grandNight,  styles.header);
        setNumCell(totalRow, 9,  grandHol,    styles.header);
        setNumCell(totalRow, 10, grandHolExt, styles.header);
    }

    // ─────────────────────────────────────────────────────────
    // 스타일
    // ─────────────────────────────────────────────────────────
    private static class StyleSet {
        final CellStyle title, header, data;

        StyleSet(XSSFWorkbook wb) {
            title = wb.createCellStyle();
            XSSFFont tf = wb.createFont();
            tf.setBold(true); tf.setFontHeightInPoints((short) 14);
            title.setFont(tf);
            title.setAlignment(HorizontalAlignment.CENTER);

            header = wb.createCellStyle();
            XSSFFont hf = wb.createFont();
            hf.setBold(true);
            header.setFont(hf);
            header.setAlignment(HorizontalAlignment.CENTER);
            header.setVerticalAlignment(VerticalAlignment.CENTER);
            setBorder(header);
            ((XSSFCellStyle) header).setFillForegroundColor(
                    new XSSFColor(new byte[]{(byte)217, (byte)217, (byte)217}, null));
            header.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            data = wb.createCellStyle();
            data.setAlignment(HorizontalAlignment.CENTER);
            data.setVerticalAlignment(VerticalAlignment.CENTER);
            setBorder(data);
        }

        private void setBorder(CellStyle s) {
            s.setBorderTop(BorderStyle.THIN);
            s.setBorderBottom(BorderStyle.THIN);
            s.setBorderLeft(BorderStyle.THIN);
            s.setBorderRight(BorderStyle.THIN);
        }
    }

    // ─────────────────────────────────────────────────────────
    // 유틸
    // ─────────────────────────────────────────────────────────
    private void setCell(Row row, int col, String value, CellStyle style) {
        Cell cell = row.createCell(col);
        cell.setCellValue(value != null ? value : "");
        cell.setCellStyle(style);
    }

    private void setNumCell(Row row, int col, Double value, CellStyle style) {
        Cell cell = row.createCell(col);
        cell.setCellValue(value != null ? value : 0.0);
        cell.setCellStyle(style);
    }

    private Row getOrCreateRow(XSSFSheet sheet, int rowIdx) {
        Row row = sheet.getRow(rowIdx);
        return row != null ? row : sheet.createRow(rowIdx);
    }

    private double nvl(Double v) { return v != null ? v : 0.0; }

    private String prevMonth(String applyYearMonth) {
        int month = Integer.parseInt(applyYearMonth.substring(5));
        return month == 1 ? "12" : String.valueOf(month - 1);
    }
}