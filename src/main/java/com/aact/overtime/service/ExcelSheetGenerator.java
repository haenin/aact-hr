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

@Component
@RequiredArgsConstructor
public class ExcelSheetGenerator {

    private final OvertimeRecordRepository overtimeRecordRepository;
    private final DepartmentSummaryRepository departmentSummaryRepository;

    private static final int COL_DEPT   = (int)(8.75  * 256);
    private static final int COL_MONTH  = (int)(9.875 * 256);
    private static final int COL_REASON = (int)(23.625 * 256);

    private static final int COL_D_DEPT  = 3616;
    private static final int COL_D_NAME  = 3648;
    private static final int COL_D_DATE  = 3552;
    private static final int COL_D_TIME  = 2304;
    private static final int COL_D_HOURS = 3648;

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

            List<String> effectiveApprovers = (approvers != null && !approvers.isEmpty())
                    ? approvers
                    : Arrays.asList("담 당", "과 장", "부 장");

            for (String dept : departments) {
                createCoverSheet(wb, styles, applyYearMonth, dept,
                        summaryByDept.getOrDefault(dept, Collections.emptyList()),
                        effectiveApprovers);

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
        int month     = Integer.parseInt(applyYearMonth.substring(5));
        int prevMonth = (month == 1) ? 12 : month - 1;

        XSSFSheet sheet = wb.createSheet(dept + "표지");

        sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
        sheet.getPrintSetup().setLandscape(false);
        sheet.setFitToPage(true);
        sheet.getPrintSetup().setFitWidth((short) 1);
        sheet.getPrintSetup().setFitHeight((short) 0);

        sheet.setMargin(Sheet.LeftMargin,   0.1968503937007874);
        sheet.setMargin(Sheet.RightMargin,  0.15748031496062992);
        sheet.setMargin(Sheet.TopMargin,    0.7480314960629921);
        sheet.setMargin(Sheet.BottomMargin, 0.7480314960629921);
        sheet.setMargin(Sheet.HeaderMargin, 0.31496062992125984);
        sheet.setMargin(Sheet.FooterMargin, 0.31496062992125984);

        sheet.setColumnWidth(0,  COL_DEPT);
        sheet.setColumnWidth(1,  COL_MONTH);
        sheet.setColumnWidth(2,  COL_MONTH);
        sheet.setColumnWidth(3,  COL_MONTH);
        sheet.setColumnWidth(4,  COL_MONTH);
        sheet.setColumnWidth(5,  COL_MONTH);
        sheet.setColumnWidth(6,  COL_MONTH);
        sheet.setColumnWidth(7,  COL_MONTH);
        sheet.setColumnWidth(8,  COL_MONTH);
        sheet.setColumnWidth(9,  COL_MONTH);
        sheet.setColumnWidth(10, COL_REASON);

        // ── row 0: 제목
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 10));
        Row r0 = sheet.createRow(0);
        r0.setHeightInPoints(27.0f);
        Cell titleCell = r0.createCell(0);
        titleCell.setCellValue("시간 외 근무수당 신청서(" + month + "월)");
        titleCell.setCellStyle(styles.coverTitle);

        // ── row 1: 여백
        sheet.createRow(1).setHeightInPoints(9.0f);

        // ── row 2~5: 결재란
        {
            int approverCount = approvers.size();
            int totalCols  = approverCount + 1;
            int colEnd     = 10;
            int colStart   = colEnd - totalCols + 1;

            for (int i = 2; i <= 5; i++) {
                sheet.createRow(i).setHeightInPoints(18.75f);
            }

            sheet.addMergedRegion(new CellRangeAddress(2, 5, colStart, colStart));
            fillMergedCells(sheet, wb, 2, 5, colStart, colStart, "결\n재", styles.approvalLabel);

            for (int i = 0; i < approverCount; i++) {
                int c = colStart + 1 + i;
                if (c > colEnd) c = colEnd;

                fillMergedCells(sheet, wb, 2, 2, c, c, approvers.get(i), styles.approvalHeader);
                sheet.addMergedRegion(new CellRangeAddress(3, 5, c, c));
                fillMergedCells(sheet, wb, 3, 5, c, c, "", styles.approvalBody);
            }
        }

        // ── row 6: 여백
        sheet.createRow(6).setHeightInPoints(9.0f);

        // ── row 7: 1. 부서별 집계표
        Row r7 = sheet.createRow(7);
        r7.setHeightInPoints(18.75f);
        Cell sectionCell = r7.createCell(0);
        sectionCell.setCellValue("1. 부서별 집계표");
        sectionCell.setCellStyle(styles.sectionTitle);

        // ── row 8: 여백
        sheet.createRow(8).setHeightInPoints(9.0f);

        // ── row 9~10: 헤더
        sheet.createRow(9).setHeightInPoints(20.25f);
        sheet.createRow(10).setHeightInPoints(24.0f);

        sheet.addMergedRegion(new CellRangeAddress(9, 10, 0, 0));
        sheet.addMergedRegion(new CellRangeAddress(9, 9,  1,  3));
        sheet.addMergedRegion(new CellRangeAddress(9, 9,  4,  6));
        sheet.addMergedRegion(new CellRangeAddress(9, 9,  7,  9));
        sheet.addMergedRegion(new CellRangeAddress(9, 10, 10, 10));

        fillMergedCells(sheet, wb, 9, 10, 0,  0,  "부서명",                    styles.header);
        fillMergedCells(sheet, wb, 9, 9,  1,  3,  "전월(" + prevMonth + "월)", styles.header);

        // 당월 헤더 1행 (col 4~6 병합): top=medium, left(col4)=medium, right(col6)=medium, bottom=thin
        {
            Row r9 = sheet.getRow(9);
            // col4: value 있음, top/left=medium, bottom/right=thin
            setCellCustomBorder(r9, 4, "당월(" + month + "월)", styles.header, wb,
                    BorderStyle.MEDIUM, BorderStyle.THIN, BorderStyle.MEDIUM, BorderStyle.THIN);
            // col5: top=medium, bottom=thin, left/right=none
            setCellCustomBorder(r9, 5, null, styles.header, wb,
                    BorderStyle.MEDIUM, BorderStyle.THIN, BorderStyle.NONE, BorderStyle.NONE);
            // col6: top=medium, bottom=thin, left=none, right=medium
            setCellCustomBorder(r9, 6, null, styles.header, wb,
                    BorderStyle.MEDIUM, BorderStyle.THIN, BorderStyle.NONE, BorderStyle.MEDIUM);
        }

        fillMergedCells(sheet, wb, 9, 9,  7,  9,  "증감",    styles.header);
        fillMergedCells(sheet, wb, 9, 10, 10, 10, "증감사유", styles.header);

        // ── row 10: 헤더 2행 (연장/야간/휴일 × 3)
        Row r10 = sheet.getRow(10);
        setCell(r10, 1, "연장", styles.header);
        setCell(r10, 2, "야간", styles.header);
        setCell(r10, 3, "휴일", styles.header);
        // 당월 2행: col4(연장) left=medium, col5(야간) thin, col6(휴일) right=medium
        setCellOuterBorder(r10, 4, "연장", styles.header, wb,
                BorderStyle.THIN, BorderStyle.THIN, BorderStyle.MEDIUM, BorderStyle.THIN);
        setCell(r10, 5, "야간", styles.header);
        setCellOuterBorder(r10, 6, "휴일", styles.header, wb,
                BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.MEDIUM);
        setCell(r10, 7, "연장", styles.header);
        setCell(r10, 8, "야간", styles.header);
        setCell(r10, 9, "휴일", styles.header);

        // ── row 11~: 데이터 행
        int rowIdx = 11;
        double totalPrevExt = 0, totalPrevNight = 0, totalPrevHol = 0;
        double totalCurrExt = 0, totalCurrNight = 0, totalCurrHol = 0;

        for (DepartmentSummary s : summaries) {
            Row dr = sheet.createRow(rowIdx);
            dr.setHeightInPoints(56.25f);

            setCell   (dr, 0,  s.getDepartment(),                                      styles.dataCenter);
            setNumCell(dr, 1,  s.getPrevExtensionHours(),                              styles.dataCenter);
            setNumCell(dr, 2,  s.getPrevNightHours(),                                  styles.dataCenter);
            setNumCell(dr, 3,  s.getPrevHolidayHours(),                                styles.dataCenter);
            // 당월 데이터: col4 left=medium, col5 thin, col6 right=medium
            setNumCellOuterBorder(dr, 4, s.getCurrExtensionHours(), styles.dataCenter, wb,
                    BorderStyle.THIN, BorderStyle.THIN, BorderStyle.MEDIUM, BorderStyle.THIN);
            setNumCell(dr, 5,  s.getCurrNightHours(),                                  styles.dataCenter);
            setNumCellOuterBorder(dr, 6, s.getCurrHolidayHours(),   styles.dataCenter, wb,
                    BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.MEDIUM);
            setNumCell(dr, 7,  s.getDiffExtensionHours(),                              styles.dataCenter);
            setNumCell(dr, 8,  s.getDiffNightHours(),                                  styles.dataCenter);
            setNumCell(dr, 9,  s.getDiffHolidayHours(),                                styles.dataCenter);
            setCell   (dr, 10, s.getChangeReason() != null ? s.getChangeReason() : "", styles.dataCenter);

            totalPrevExt   += nvl(s.getPrevExtensionHours());
            totalPrevNight += nvl(s.getPrevNightHours());
            totalPrevHol   += nvl(s.getPrevHolidayHours());
            totalCurrExt   += nvl(s.getCurrExtensionHours());
            totalCurrNight += nvl(s.getCurrNightHours());
            totalCurrHol   += nvl(s.getCurrHolidayHours());
            rowIdx++;
        }

        // ── 총 합계 행: col4 left=medium/bottom=medium, col5 bottom=medium, col6 right/bottom=medium
        Row totalRow = sheet.createRow(rowIdx);
        totalRow.setHeightInPoints(56.25f);
        setCell   (totalRow, 0,  "총 합계",                      styles.header);
        setNumCell(totalRow, 1,  totalPrevExt,                   styles.header);
        setNumCell(totalRow, 2,  totalPrevNight,                 styles.header);
        setNumCell(totalRow, 3,  totalPrevHol,                   styles.header);
        // 당월 합계: col4 top/left/bottom=medium, right=thin
        setNumCellOuterBorder(totalRow, 4, totalCurrExt,   styles.header, wb,
                BorderStyle.THIN, BorderStyle.MEDIUM, BorderStyle.MEDIUM, BorderStyle.THIN);
        // col5 top=thin, bottom=medium, left/right=thin
        setNumCellOuterBorder(totalRow, 5, totalCurrNight, styles.header, wb,
                BorderStyle.THIN, BorderStyle.MEDIUM, BorderStyle.THIN, BorderStyle.THIN);
        // col6 top=thin, bottom/right=medium, left=thin
        setNumCellOuterBorder(totalRow, 6, totalCurrHol,   styles.header, wb,
                BorderStyle.THIN, BorderStyle.MEDIUM, BorderStyle.THIN, BorderStyle.MEDIUM);
        setNumCell(totalRow, 7,  totalCurrExt   - totalPrevExt,  styles.header);
        setNumCell(totalRow, 8,  totalCurrNight - totalPrevNight, styles.header);
        setNumCell(totalRow, 9,  totalCurrHol   - totalPrevHol,  styles.header);
        setCell   (totalRow, 10, "",                             styles.header);
        rowIdx++;

        // ── 2행 빈행 후 "2. 부서별 증감 사유"
        sheet.createRow(rowIdx++).setHeightInPoints(18.75f);
        sheet.createRow(rowIdx++).setHeightInPoints(18.75f);

        Row r2s = sheet.createRow(rowIdx++);
        r2s.setHeightInPoints(18.75f);
        r2s.createCell(0).setCellValue("2. 부서별 증감 사유 : 붙임참조");
        r2s.getCell(0).setCellStyle(styles.sectionTitle);

        sheet.createRow(rowIdx++).setHeightInPoints(18.75f);
        sheet.createRow(rowIdx++).setHeightInPoints(18.75f);

        // ── 3. 개인별 집계표
        Row r3s = sheet.createRow(rowIdx++);
        r3s.setHeightInPoints(18.75f);
        r3s.createCell(0).setCellValue("3. 개인별 집계표 : 붙임참조");
        r3s.getCell(0).setCellStyle(styles.sectionTitle);

        sheet.createRow(rowIdx++).setHeightInPoints(18.75f);

        // ── 안내 문구
        sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx, 0, 10));
        Row noteRow = sheet.createRow(rowIdx++);
        Cell noteCell = noteRow.createCell(0);
        noteCell.setCellValue("\u3000\u3000상기와 같이 " + month + "월 시간외 수당을 신청하오니 검토하시고 재가하여 주시기 바랍니다.");
        noteCell.setCellStyle(styles.noteText);

        sheet.createRow(rowIdx++).setHeightInPoints(250.75f);
        // ── 날짜
        rowIdx += 5;
        sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx, 0, 10));
        Row dateRow = sheet.createRow(rowIdx++);
        dateRow.createCell(0).setCellValue(
                applyYearMonth.substring(0, 4) + ". " + String.format("%02d", month) + ". 07");
        dateRow.getCell(0).setCellStyle(styles.dateText);

        // ── 회사명
        rowIdx += 2;
        sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx + 1, 0, 10));
        Row compRow = sheet.createRow(rowIdx);
        compRow.setHeightInPoints(25.0f);
        Cell compCell = compRow.createCell(0);
        compCell.setCellValue("에 이 에 이 씨 티 유 한 회 사");
        compCell.setCellStyle(styles.companyName);
    }

    // ─────────────────────────────────────────────────────────
    // 세부내역 시트
    // ─────────────────────────────────────────────────────────
    private void createDetailSheet(XSSFWorkbook wb, StyleSet styles,
                                   String applyYearMonth, String dept,
                                   List<OvertimeRecord> records) {
        int month = Integer.parseInt(applyYearMonth.substring(5));
        XSSFSheet sheet = wb.createSheet(dept + "(" + month + "월)");

        sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
        sheet.setFitToPage(true);
        sheet.getPrintSetup().setFitWidth((short) 1);
        sheet.getPrintSetup().setFitHeight((short) 0);

        sheet.setColumnWidth(0,  COL_D_DEPT);
        sheet.setColumnWidth(1,  COL_D_NAME);
        sheet.setColumnWidth(2,  COL_D_DATE);
        sheet.setColumnWidth(3,  COL_D_TIME);
        sheet.setColumnWidth(4,  COL_D_TIME);
        sheet.setColumnWidth(5,  COL_D_TIME);
        sheet.setColumnWidth(6,  COL_D_TIME);
        sheet.setColumnWidth(7,  COL_D_HOURS);
        sheet.setColumnWidth(8,  COL_D_HOURS);
        sheet.setColumnWidth(9,  COL_D_HOURS);
        sheet.setColumnWidth(10, COL_D_HOURS);

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 10));
        Row r0 = sheet.createRow(0);
        setRowHeight(sheet, 0, 26.25f);
        Cell titleCell = r0.createCell(0);
        titleCell.setCellValue(month + "월 시간 외 근로시간 개인별 세부내역 (" + dept + ")");
        titleCell.setCellStyle(styles.detailTitle);

        setRowHeight(sheet, 1, 8.0f);

        sheet.createRow(2);
        sheet.createRow(3);
        setRowHeight(sheet, 2, 20.0f);
        setRowHeight(sheet, 3, 20.0f);

        sheet.addMergedRegion(new CellRangeAddress(2, 3, 0, 0));
        fillMergedCells(sheet, wb, 2, 3, 0, 0, "부서", styles.header);
        sheet.addMergedRegion(new CellRangeAddress(2, 3, 1, 1));
        fillMergedCells(sheet, wb, 2, 3, 1, 1, "성명", styles.header);
        sheet.addMergedRegion(new CellRangeAddress(2, 3, 2, 2));
        fillMergedCells(sheet, wb, 2, 3, 2, 2, "일자", styles.header);
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 3, 4));
        fillMergedCells(sheet, wb, 2, 2, 3, 4, "예정근무", styles.header);
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 5, 6));
        fillMergedCells(sheet, wb, 2, 2, 5, 6, "실근무(지문)", styles.header);
        sheet.addMergedRegion(new CellRangeAddress(2, 3, 7, 7));
        fillMergedCells(sheet, wb, 2, 3, 7, 7, "연장", styles.header);
        sheet.addMergedRegion(new CellRangeAddress(2, 3, 8, 8));
        fillMergedCells(sheet, wb, 2, 3, 8, 8, "야간", styles.header);
        sheet.addMergedRegion(new CellRangeAddress(2, 3, 9, 9));
        fillMergedCells(sheet, wb, 2, 3, 9, 9, "휴일", styles.header);
        sheet.addMergedRegion(new CellRangeAddress(2, 3, 10, 10));
        fillMergedCells(sheet, wb, 2, 3, 10, 10, "휴일연장", styles.header);

        Row r3 = sheet.getRow(3);
        setCell(r3, 3, "출근", styles.header);
        setCell(r3, 4, "퇴근", styles.header);
        setCell(r3, 5, "출근", styles.header);
        setCell(r3, 6, "퇴근", styles.header);

        Map<String, List<OvertimeRecord>> byEmployee = new LinkedHashMap<>();
        for (OvertimeRecord r : records) {
            byEmployee.computeIfAbsent(r.getEmployeeName(), k -> new ArrayList<>()).add(r);
        }

        int rowIdx    = 4;
        int deptStart = rowIdx;
        double grandExt = 0, grandNight = 0, grandHol = 0, grandHolExt = 0;

        for (Map.Entry<String, List<OvertimeRecord>> entry : byEmployee.entrySet()) {
            String empName = entry.getKey();
            List<OvertimeRecord> empRecords = entry.getValue();
            double subExt = 0, subNight = 0, subHol = 0, subHolExt = 0;
            int empStart = rowIdx;

            for (OvertimeRecord rec : empRecords) {
                Row dr = sheet.createRow(rowIdx++);
                setRowHeight(sheet, rowIdx - 1, 19.5f);
                dr.createCell(0).setCellStyle(styles.data);
                setCell(dr, 2,  rec.getWorkDate()       != null ? rec.getWorkDate().toString()       : "", styles.data);
                setCell(dr, 3,  rec.getScheduledStart() != null ? rec.getScheduledStart().toString() : "", styles.data);
                setCell(dr, 4,  rec.getScheduledEnd()   != null ? rec.getScheduledEnd().toString()   : "", styles.data);
                setCell(dr, 5,  rec.getActualStart()    != null ? rec.getActualStart().toString()    : "", styles.data);
                setCell(dr, 6,  rec.getActualEnd()      != null ? rec.getActualEnd().toString()      : "", styles.data);
                setNumCell(dr, 7,  rec.getExtensionHours(),        styles.data);
                setNumCell(dr, 8,  rec.getNightHours(),            styles.data);
                setNumCell(dr, 9,  rec.getHolidayHours(),          styles.data);
                setNumCell(dr, 10, rec.getHolidayExtensionHours(), styles.data);

                subExt    += nvl(rec.getExtensionHours());
                subNight  += nvl(rec.getNightHours());
                subHol    += nvl(rec.getHolidayHours());
                subHolExt += nvl(rec.getHolidayExtensionHours());
            }

            if (empRecords.size() > 1) {
                sheet.addMergedRegion(new CellRangeAddress(empStart, rowIdx - 1, 1, 1));
            }
            setCell(sheet.getRow(empStart), 1, empName, styles.data);

            Row subRow = sheet.createRow(rowIdx);
            setRowHeight(sheet, rowIdx, 19.5f);
            subRow.createCell(0).setCellStyle(styles.data);
            sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx, 1, 6));
            fillMergedCells(sheet, wb, rowIdx, rowIdx, 1, 6, "소계", styles.header);
            setNumCell(subRow, 7,  subExt,    styles.header);
            setNumCell(subRow, 8,  subNight,  styles.header);
            setNumCell(subRow, 9,  subHol,    styles.header);
            setNumCell(subRow, 10, subHolExt, styles.header);
            rowIdx++;

            grandExt    += subExt;
            grandNight  += subNight;
            grandHol    += subHol;
            grandHolExt += subHolExt;
        }

        if (rowIdx - 1 > deptStart) {
            sheet.addMergedRegion(new CellRangeAddress(deptStart, rowIdx - 1, 0, 0));
        }
        setCell(sheet.getRow(deptStart), 0, records.get(0).getDepartment(), styles.data);

        Row totalRow = sheet.createRow(rowIdx);
        setRowHeight(sheet, rowIdx, 24.75f);
        sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx, 0, 6));
        fillMergedCells(sheet, wb, rowIdx, rowIdx, 0, 6, dept + " 합계", styles.header);
        setNumCell(totalRow, 7,  grandExt,    styles.header);
        setNumCell(totalRow, 8,  grandNight,  styles.header);
        setNumCell(totalRow, 9,  grandHol,    styles.header);
        setNumCell(totalRow, 10, grandHolExt, styles.header);
    }

    // ─────────────────────────────────────────────────────────
    // StyleSet
    // ─────────────────────────────────────────────────────────
    private static class StyleSet {
        final CellStyle coverTitle;
        final CellStyle sectionTitle;
        final CellStyle detailTitle;
        final CellStyle header;
        final CellStyle data;
        final CellStyle dataCenter;
        final CellStyle dataLeft;
        final CellStyle approvalLabel;
        final CellStyle approvalHeader;
        final CellStyle approvalBody;
        final CellStyle noteText;
        final CellStyle dateText;
        final CellStyle companyName;

        StyleSet(XSSFWorkbook wb) {
            byte[] gray = {(byte)217, (byte)217, (byte)217};

            coverTitle = wb.createCellStyle();
            XSSFFont cf = wb.createFont();
            cf.setFontName("HY헤드라인M"); cf.setFontHeightInPoints((short)22); cf.setBold(false);
            coverTitle.setFont(cf);
            coverTitle.setAlignment(HorizontalAlignment.CENTER);
            coverTitle.setVerticalAlignment(VerticalAlignment.CENTER);

            sectionTitle = wb.createCellStyle();
            XSSFFont sf = wb.createFont();
            sf.setFontName("돋움"); sf.setFontHeightInPoints((short)14); sf.setBold(true);
            sectionTitle.setFont(sf);
            sectionTitle.setVerticalAlignment(VerticalAlignment.CENTER);

            detailTitle = wb.createCellStyle();
            XSSFFont dtf = wb.createFont();
            dtf.setFontName("돋움"); dtf.setFontHeightInPoints((short)14); dtf.setBold(false);
            detailTitle.setFont(dtf);
            detailTitle.setAlignment(HorizontalAlignment.CENTER);
            detailTitle.setVerticalAlignment(VerticalAlignment.CENTER);

            header = wb.createCellStyle();
            XSSFFont hf = wb.createFont();
            hf.setFontName("돋움"); hf.setFontHeightInPoints((short)12); hf.setBold(true);
            header.setFont(hf);
            header.setAlignment(HorizontalAlignment.CENTER);
            header.setVerticalAlignment(VerticalAlignment.CENTER);
            setThin(header);
            ((XSSFCellStyle)header).setFillForegroundColor(new XSSFColor(gray, null));
            header.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            data = wb.createCellStyle();
            data.setAlignment(HorizontalAlignment.CENTER);
            data.setVerticalAlignment(VerticalAlignment.CENTER);
            setThin(data);

            dataCenter = wb.createCellStyle();
            dataCenter.cloneStyleFrom(data);

            dataLeft = wb.createCellStyle();
            dataLeft.cloneStyleFrom(data);
            dataLeft.setAlignment(HorizontalAlignment.LEFT);

            approvalLabel = wb.createCellStyle();
            XSSFFont alf = wb.createFont();
            alf.setFontName("돋움"); alf.setFontHeightInPoints((short)11);
            approvalLabel.setFont(alf);
            approvalLabel.setAlignment(HorizontalAlignment.CENTER);
            approvalLabel.setVerticalAlignment(VerticalAlignment.CENTER);
            approvalLabel.setWrapText(true);
            setThin(approvalLabel);

            approvalHeader = wb.createCellStyle();
            XSSFFont ahf = wb.createFont();
            ahf.setFontName("돋움"); ahf.setFontHeightInPoints((short)11);
            approvalHeader.setFont(ahf);
            approvalHeader.setAlignment(HorizontalAlignment.CENTER);
            approvalHeader.setVerticalAlignment(VerticalAlignment.CENTER);
            setThin(approvalHeader);

            approvalBody = wb.createCellStyle();
            approvalBody.setAlignment(HorizontalAlignment.CENTER);
            approvalBody.setVerticalAlignment(VerticalAlignment.CENTER);
            setThin(approvalBody);

            noteText = wb.createCellStyle();
            XSSFFont nf = wb.createFont();
            nf.setFontName("돋움"); nf.setFontHeightInPoints((short)12);
            noteText.setFont(nf);
            noteText.setVerticalAlignment(VerticalAlignment.CENTER);

            dateText = wb.createCellStyle();
            XSSFFont datef = wb.createFont();
            datef.setFontName("돋움"); datef.setFontHeightInPoints((short)14);
            dateText.setFont(datef);
            dateText.setAlignment(HorizontalAlignment.CENTER);
            dateText.setVerticalAlignment(VerticalAlignment.CENTER);

            companyName = wb.createCellStyle();
            XSSFFont cnf = wb.createFont();
            cnf.setFontName("돋움"); cnf.setFontHeightInPoints((short)18); cnf.setBold(true);
            companyName.setFont(cnf);
            companyName.setAlignment(HorizontalAlignment.CENTER);
            companyName.setVerticalAlignment(VerticalAlignment.CENTER);
        }

        private void setThin(CellStyle s) {
            s.setBorderTop(BorderStyle.THIN);
            s.setBorderBottom(BorderStyle.THIN);
            s.setBorderLeft(BorderStyle.THIN);
            s.setBorderRight(BorderStyle.THIN);
        }
    }

    // ─────────────────────────────────────────────────────────
    // 커스텀 border 헬퍼 (문자열, null이면 값 안 씀)
    // top / bottom / left / right 순
    // ─────────────────────────────────────────────────────────
    private void setCellCustomBorder(Row row, int col, String value, CellStyle baseStyle,
                                     XSSFWorkbook wb,
                                     BorderStyle top, BorderStyle bottom,
                                     BorderStyle left, BorderStyle right) {
        Cell cell = row.getCell(col);
        if (cell == null) cell = row.createCell(col);
        if (value != null) cell.setCellValue(value);
        XSSFCellStyle cs = wb.createCellStyle();
        cs.cloneStyleFrom(baseStyle);
        cs.setBorderTop(top);
        cs.setBorderBottom(bottom);
        cs.setBorderLeft(left);
        cs.setBorderRight(right);
        cell.setCellStyle(cs);
    }

    // ─────────────────────────────────────────────────────────
    // 커스텀 border 헬퍼 (문자열, 기존 setCellOuterBorder 유지)
    // ─────────────────────────────────────────────────────────
    private void setCellOuterBorder(Row row, int col, String value, CellStyle baseStyle,
                                    XSSFWorkbook wb,
                                    BorderStyle top, BorderStyle bottom,
                                    BorderStyle left, BorderStyle right) {
        Cell cell = row.createCell(col);
        if (value != null) cell.setCellValue(value);
        XSSFCellStyle cs = wb.createCellStyle();
        cs.cloneStyleFrom(baseStyle);
        cs.setBorderTop(top);
        cs.setBorderBottom(bottom);
        cs.setBorderLeft(left);
        cs.setBorderRight(right);
        cell.setCellStyle(cs);
    }

    // ─────────────────────────────────────────────────────────
    // 커스텀 border 헬퍼 (숫자)
    // ─────────────────────────────────────────────────────────
    private void setNumCellOuterBorder(Row row, int col, Double value, CellStyle baseStyle,
                                       XSSFWorkbook wb,
                                       BorderStyle top, BorderStyle bottom,
                                       BorderStyle left, BorderStyle right) {
        Cell cell = row.createCell(col);
        cell.setCellValue(value != null ? value : 0.0);
        XSSFCellStyle cs = wb.createCellStyle();
        cs.cloneStyleFrom(baseStyle);
        cs.setBorderTop(top);
        cs.setBorderBottom(bottom);
        cs.setBorderLeft(left);
        cs.setBorderRight(right);
        cell.setCellStyle(cs);
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

    private void setRowHeight(XSSFSheet sheet, int rowIdx, float height) {
        getOrCreateRow(sheet, rowIdx).setHeightInPoints(height);
    }

    private double nvl(Double v) { return v != null ? v : 0.0; }

    private void fillMergedCells(XSSFSheet sheet, XSSFWorkbook wb,
                                 int firstRow, int lastRow,
                                 int firstCol, int lastCol,
                                 String value, CellStyle baseStyle) {
        for (int r = firstRow; r <= lastRow; r++) {
            Row row = getOrCreateRow(sheet, r);
            for (int c = firstCol; c <= lastCol; c++) {
                Cell cell = row.getCell(c);
                if (cell == null) cell = row.createCell(c);
                if (r == firstRow && c == firstCol) {
                    cell.setCellValue(value != null ? value : "");
                }
                XSSFCellStyle cs = wb.createCellStyle();
                cs.cloneStyleFrom(baseStyle);
                cs.setBorderTop(BorderStyle.THIN);
                cs.setBorderBottom(BorderStyle.THIN);
                cs.setBorderLeft(BorderStyle.THIN);
                cs.setBorderRight(BorderStyle.THIN);
                cell.setCellStyle(cs);
            }
        }
    }
}