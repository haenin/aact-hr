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
 * A4 용지 기준으로 인쇄 영역에 맞게 설계
 * 결재란은 approvers 리스트로 동적 처리
 */
@Component
@RequiredArgsConstructor
public class ExcelSheetGenerator {

    private final OvertimeRecordRepository overtimeRecordRepository;
    private final DepartmentSummaryRepository departmentSummaryRepository;

    // ── A4 기준 컬럼 너비 (1/256 단위)
    // 표지: 부서명(A) + 전월/당월/증감 각 3열 + 증감사유
    // 컬럼 인덱스: A=0, B=1, C=2, D=3, E=4, F=5, G=6, H=7, I=8, J=9, K=10
    private static final int COL_DEPT    = 3584; // 14.0
    private static final int COL_MONTH   = 2048; // 8.0  (전월/당월/증감 각 컬럼)
    private static final int COL_REASON  = 6400; // 25.0 (증감사유)

    // 세부내역 컬럼 너비
    private static final int COL_D_DEPT  = 3616; // 14.125 부서
    private static final int COL_D_NAME  = 3648; // 14.25  성명
    private static final int COL_D_DATE  = 3552; // 13.875 일자
    private static final int COL_D_TIME  = 2304; // 9.0    출퇴근시간
    private static final int COL_D_HOURS = 3648; // 14.25  근무시간

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
    //
    // A4 컬럼 구성 (0-based):
    //  0  : 부서명
    //  1  : 전월 연장
    //  2  : 전월 야간
    //  3  : 전월 휴일
    //  4  : 당월 연장
    //  5  : 당월 야간
    //  6  : 당월 휴일
    //  7  : 증감 연장
    //  8  : 증감 야간
    //  9  : 증감 휴일
    //  10 : 증감사유
    //
    // 결재란: 우측 상단 (row 2~5)
    //  ┌────┬──────┬──────┬──────┬──────┐
    //  │결재│ 담 당│ 과 장│ 상 무│부사장│
    //  │    │      │      │      │      │
    //  └────┴──────┴──────┴──────┴──────┘
    // ─────────────────────────────────────────────────────────
    private void createCoverSheet(XSSFWorkbook wb, StyleSet styles,
                                  String applyYearMonth, String dept,
                                  List<DepartmentSummary> summaries,
                                  List<String> approvers) {
        int month     = Integer.parseInt(applyYearMonth.substring(5));
        int prevMonth = month == 1 ? 12 : month - 1;

        XSSFSheet sheet = wb.createSheet(dept + "표지");

        // 인쇄 설정 (A4)
        sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
        sheet.getPrintSetup().setLandscape(false);
        sheet.setFitToPage(true);
        sheet.getPrintSetup().setFitWidth((short) 1);
        sheet.getPrintSetup().setFitHeight((short) 0);

        // ── 열 너비 설정
        sheet.setColumnWidth(0,  COL_DEPT);   // 부서명
        sheet.setColumnWidth(1,  COL_MONTH);  // 전월 연장
        sheet.setColumnWidth(2,  COL_MONTH);  // 전월 야간
        sheet.setColumnWidth(3,  COL_MONTH);  // 전월 휴일
        sheet.setColumnWidth(4,  COL_MONTH);  // 당월 연장
        sheet.setColumnWidth(5,  COL_MONTH);  // 당월 야간
        sheet.setColumnWidth(6,  COL_MONTH);  // 당월 휴일
        sheet.setColumnWidth(7,  COL_MONTH);  // 증감 연장
        sheet.setColumnWidth(8,  COL_MONTH);  // 증감 야간
        sheet.setColumnWidth(9,  COL_MONTH);  // 증감 휴일
        sheet.setColumnWidth(10, COL_REASON); // 증감사유

        // ── row 0: 제목 (A1:K1 병합, 볼드X, 22pt)
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 10));
        Row r0 = sheet.createRow(0);
        setRowHeight(sheet, 0, 30.0f);
        Cell titleCell = r0.createCell(0);
        titleCell.setCellValue("시간 외 근무수당 신청서(" + month + "월)");
        titleCell.setCellStyle(styles.coverTitle);

        // ── row 1: 여백
        setRowHeight(sheet, 1, 10.0f);

        // ── row 2~5: 결재란 (우측)
        // 결재란: col 6~10 구간 사용
        // ┌────┬──────┬──────┬──────┬──────┐
        // │결재│ 담 당│ 과 장│ 상 무│부사장│  row 2
        // │    │      │      │      │      │  row 3~5 (서명칸)
        // └────┴──────┴──────┴──────┴──────┘
        if (!approvers.isEmpty()) {
            setRowHeight(sheet, 2, 18.0f);
            setRowHeight(sheet, 3, 25.0f);
            setRowHeight(sheet, 4, 25.0f);
            setRowHeight(sheet, 5, 25.0f);

            // "결재" 라벨 (col 6, row 2~5 병합)
            sheet.addMergedRegion(new CellRangeAddress(2, 5, 6, 6));
            fillMergedCells(sheet, wb, 2, 5, 6, 6, "결\n재", styles.approvalLabel);

            // 결재자 헤더 + 서명칸
            int totalCols   = 10 - 7 + 1; // col 7~10 = 4칸
            int colsPerApprover = totalCols / approvers.size();

            for (int i = 0; i < approvers.size(); i++) {
                int colStart = 7 + (i * colsPerApprover);
                int colEnd   = (i == approvers.size() - 1) ? 10 : colStart + colsPerApprover - 1;

                // 헤더 (row 2)
                sheet.addMergedRegion(new CellRangeAddress(2, 2, colStart, colEnd));
                fillMergedCells(sheet, wb, 2, 2, colStart, colEnd, approvers.get(i), styles.approvalHeader);

                // 서명칸 (row 3~5 하나로 병합)
                sheet.addMergedRegion(new CellRangeAddress(3, 5, colStart, colEnd));
                fillMergedCells(sheet, wb, 3, 5, colStart, colEnd, "", styles.approvalBody);
            }
        }

        // ── row 6: 1. 부서별 집계표
        setRowHeight(sheet, 6, 20.0f);
        Row r6 = getOrCreateRow(sheet, 6);
        Cell sectionCell = r6.createCell(0);
        sectionCell.setCellValue("1. 부서별 집계표");
        sectionCell.setCellStyle(styles.sectionTitle);

        // ── row 7: 헤더 1행
        // A=부서명(2행병합), B~D=전월, E~G=당월, H~J=증감, K=증감사유(2행병합)
        setRowHeight(sheet, 7, 20.0f);
        Row r7 = sheet.createRow(7);
        sheet.addMergedRegion(new CellRangeAddress(7, 8, 0, 0));   // 부서명
        sheet.addMergedRegion(new CellRangeAddress(7, 7, 1, 3));   // 전월
        sheet.addMergedRegion(new CellRangeAddress(7, 7, 4, 6));   // 당월
        sheet.addMergedRegion(new CellRangeAddress(7, 7, 7, 9));   // 증감
        sheet.addMergedRegion(new CellRangeAddress(7, 8, 10, 10)); // 증감사유

        fillMergedCells(sheet, wb, 7, 8, 0,  0,  "부서명",                    styles.header);
        fillMergedCells(sheet, wb, 7, 7, 1,  3,  "전월(" + prevMonth + "월)", styles.header);
        fillMergedCells(sheet, wb, 7, 7, 4,  6,  "당월(" + month + "월)",     styles.header);
        fillMergedCells(sheet, wb, 7, 7, 7,  9,  "증감",                      styles.header);
        fillMergedCells(sheet, wb, 7, 8, 10, 10, "증감사유",                  styles.header);

        // ── row 8: 헤더 2행 (연장/야간/휴일 반복)
        setRowHeight(sheet, 8, 20.0f);
        Row r8 = sheet.createRow(8);
        for (int sc : new int[]{1, 4, 7}) {
            setCell(r8, sc,     "연장", styles.header);
            setCell(r8, sc + 1, "야간", styles.header);
            setCell(r8, sc + 2, "휴일", styles.header);
        }

        // ── row 9~: 데이터 행
        int rowIdx = 9;
        double totalPrevExt = 0, totalPrevNight = 0, totalPrevHol = 0;
        double totalCurrExt = 0, totalCurrNight = 0, totalCurrHol = 0;

        for (DepartmentSummary s : summaries) {
            setRowHeight(sheet, rowIdx, 30.0f);
            Row dr = sheet.createRow(rowIdx++);
            setCell(dr, 0, s.getDepartment(), styles.dataCenter);
            setNumCell(dr, 1, s.getPrevExtensionHours(), styles.dataCenter);
            setNumCell(dr, 2, s.getPrevNightHours(),     styles.dataCenter);
            setNumCell(dr, 3, s.getPrevHolidayHours(),   styles.dataCenter);
            setNumCell(dr, 4, s.getCurrExtensionHours(), styles.dataCenter);
            setNumCell(dr, 5, s.getCurrNightHours(),     styles.dataCenter);
            setNumCell(dr, 6, s.getCurrHolidayHours(),   styles.dataCenter);
            setNumCell(dr, 7, s.getDiffExtensionHours(), styles.dataCenter);
            setNumCell(dr, 8, s.getDiffNightHours(),     styles.dataCenter);
            setNumCell(dr, 9, s.getDiffHolidayHours(),   styles.dataCenter);
            setCell(dr, 10, s.getChangeReason() != null ? s.getChangeReason() : "", styles.dataLeft);

            totalPrevExt   += nvl(s.getPrevExtensionHours());
            totalPrevNight += nvl(s.getPrevNightHours());
            totalPrevHol   += nvl(s.getPrevHolidayHours());
            totalCurrExt   += nvl(s.getCurrExtensionHours());
            totalCurrNight += nvl(s.getCurrNightHours());
            totalCurrHol   += nvl(s.getCurrHolidayHours());
        }

        // ── 총 합계 행
        setRowHeight(sheet, rowIdx, 25.0f);
        Row totalRow = sheet.createRow(rowIdx++);
        setCell(totalRow, 0, "총 합계", styles.header);
        setNumCell(totalRow, 1, totalPrevExt,                    styles.header);
        setNumCell(totalRow, 2, totalPrevNight,                  styles.header);
        setNumCell(totalRow, 3, totalPrevHol,                    styles.header);
        setNumCell(totalRow, 4, totalCurrExt,                    styles.header);
        setNumCell(totalRow, 5, totalCurrNight,                  styles.header);
        setNumCell(totalRow, 6, totalCurrHol,                    styles.header);
        setNumCell(totalRow, 7, totalCurrExt   - totalPrevExt,   styles.header);
        setNumCell(totalRow, 8, totalCurrNight - totalPrevNight, styles.header);
        setNumCell(totalRow, 9, totalCurrHol   - totalPrevHol,   styles.header);
        setCell(totalRow, 10, "", styles.header);

        // ── 안내 문구
        rowIdx++;
        sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx, 0, 10));
        Row noteRow = sheet.createRow(rowIdx++);
        noteRow.createCell(0).setCellValue(
                "\u3000\u3000상기와 같이 " + month + "월 시간외 수당을 신청하오니 검토하시고 재가하여 주시기 바랍니다.");

        // ── 날짜
        rowIdx += 5;
        sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx, 0, 10));
        Row dateRow = sheet.createRow(rowIdx++);
        dateRow.createCell(0).setCellValue(
                applyYearMonth.substring(0, 4) + ". " + String.format("%02d", month) + ". 07");

        // ── 회사명
        rowIdx += 2;
        sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx + 1, 0, 10));
        Row compRow = sheet.createRow(rowIdx);
        setRowHeight(sheet, rowIdx, 25.0f);
        Cell compCell = compRow.createCell(0);
        compCell.setCellValue("에 이 에 이 씨 티 유 한 회 사");
        compCell.setCellStyle(styles.coverTitle);
    }

    // ─────────────────────────────────────────────────────────
    // 세부내역 시트
    //
    // A4 컬럼 구성 (0-based):
    //  0  : 부서
    //  1  : 성명
    //  2  : 일자
    //  3  : 예정 출근
    //  4  : 예정 퇴근
    //  5  : 실 출근
    //  6  : 실 퇴근
    //  7  : 연장
    //  8  : 야간
    //  9  : 휴일
    //  10 : 휴일연장
    // ─────────────────────────────────────────────────────────
    private void createDetailSheet(XSSFWorkbook wb, StyleSet styles,
                                   String applyYearMonth, String dept,
                                   List<OvertimeRecord> records) {
        int month = Integer.parseInt(applyYearMonth.substring(5));
        XSSFSheet sheet = wb.createSheet(dept + "(" + month + "월)");

        // 인쇄 설정 (A4)
        sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
        sheet.setFitToPage(true);
        sheet.getPrintSetup().setFitWidth((short) 1);
        sheet.getPrintSetup().setFitHeight((short) 0);

        // ── 열 너비
        sheet.setColumnWidth(0,  COL_D_DEPT);  // 부서
        sheet.setColumnWidth(1,  COL_D_NAME);  // 성명
        sheet.setColumnWidth(2,  COL_D_DATE);  // 일자
        sheet.setColumnWidth(3,  COL_D_TIME);  // 예정 출근
        sheet.setColumnWidth(4,  COL_D_TIME);  // 예정 퇴근
        sheet.setColumnWidth(5,  COL_D_TIME);  // 실 출근
        sheet.setColumnWidth(6,  COL_D_TIME);  // 실 퇴근
        sheet.setColumnWidth(7,  COL_D_HOURS); // 연장
        sheet.setColumnWidth(8,  COL_D_HOURS); // 야간
        sheet.setColumnWidth(9,  COL_D_HOURS); // 휴일
        sheet.setColumnWidth(10, COL_D_HOURS); // 휴일연장

        // ── row 0: 제목 (A1:K1 병합)
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 10));
        Row r0 = sheet.createRow(0);
        setRowHeight(sheet, 0, 26.25f);
        Cell titleCell = r0.createCell(0);
        titleCell.setCellValue(month + "월 시간 외 근로시간 개인별 세부내역 (" + dept + ")");
        titleCell.setCellStyle(styles.detailTitle);

        // ── row 1: 여백
        setRowHeight(sheet, 1, 8.0f);

        // ── row 2~3: 헤더 (2행)
        sheet.createRow(2);
        sheet.createRow(3);
        setRowHeight(sheet, 2, 20.0f);
        setRowHeight(sheet, 3, 20.0f);

        // A3:A4 부서
        sheet.addMergedRegion(new CellRangeAddress(2, 3, 0, 0));
        fillMergedCells(sheet, wb, 2, 3, 0, 0, "부서", styles.header);
        // B3:B4 성명
        sheet.addMergedRegion(new CellRangeAddress(2, 3, 1, 1));
        fillMergedCells(sheet, wb, 2, 3, 1, 1, "성명", styles.header);
        // C3:C4 일자
        sheet.addMergedRegion(new CellRangeAddress(2, 3, 2, 2));
        fillMergedCells(sheet, wb, 2, 3, 2, 2, "일자", styles.header);
        // D3:E3 예정근무
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 3, 4));
        fillMergedCells(sheet, wb, 2, 2, 3, 4, "예정근무", styles.header);
        // F3:G3 실근무(지문)
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 5, 6));
        fillMergedCells(sheet, wb, 2, 2, 5, 6, "실근무(지문)", styles.header);
        // H3:H4 연장
        sheet.addMergedRegion(new CellRangeAddress(2, 3, 7, 7));
        fillMergedCells(sheet, wb, 2, 3, 7, 7, "연장", styles.header);
        // I3:I4 야간
        sheet.addMergedRegion(new CellRangeAddress(2, 3, 8, 8));
        fillMergedCells(sheet, wb, 2, 3, 8, 8, "야간", styles.header);
        // J3:J4 휴일
        sheet.addMergedRegion(new CellRangeAddress(2, 3, 9, 9));
        fillMergedCells(sheet, wb, 2, 3, 9, 9, "휴일", styles.header);
        // K3:K4 휴일연장
        sheet.addMergedRegion(new CellRangeAddress(2, 3, 10, 10));
        fillMergedCells(sheet, wb, 2, 3, 10, 10, "휴일연장", styles.header);

        // row 3: 출근/퇴근
        Row r3 = sheet.getRow(3);
        setCell(r3, 3, "출근", styles.header);
        setCell(r3, 4, "퇴근", styles.header);
        setCell(r3, 5, "출근", styles.header);
        setCell(r3, 6, "퇴근", styles.header);

        // ── row 4~: 직원별 데이터
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
                dr.createCell(0).setCellStyle(styles.data); // A열 테두리
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

            // 성명 병합
            if (empRecords.size() > 1) {
                sheet.addMergedRegion(new CellRangeAddress(empStart, rowIdx - 1, 1, 1));
            }
            setCell(sheet.getRow(empStart), 1, empName, styles.data);

            // 소계 행 (B:G 병합 + 전체 테두리)
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

        // 부서 병합 (소계행 포함, 합계행 제외)
        if (rowIdx - 1 > deptStart) {
            sheet.addMergedRegion(new CellRangeAddress(deptStart, rowIdx - 1, 0, 0));
        }
        setCell(sheet.getRow(deptStart), 0, records.get(0).getDepartment(), styles.data);

        // 합계 행 (A:G 병합)
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
    // 스타일
    // ─────────────────────────────────────────────────────────
    private static class StyleSet {
        final CellStyle coverTitle;    // 표지 제목 (볼드X, 22pt, 중앙)
        final CellStyle sectionTitle;  // 섹션 제목 (볼드, 14pt)
        final CellStyle detailTitle;   // 세부내역 제목 (볼드X, 14pt, 중앙)
        final CellStyle header;        // 헤더 (볼드, 12pt, 회색배경, 테두리)
        final CellStyle data;          // 데이터 (테두리, 중앙)
        final CellStyle dataCenter;    // 데이터 중앙정렬
        final CellStyle dataLeft;      // 데이터 좌측정렬 (증감사유)
        final CellStyle approvalLabel; // 결재 라벨
        final CellStyle approvalHeader;// 결재자 직급
        final CellStyle approvalBody;  // 결재 서명칸

        StyleSet(XSSFWorkbook wb) {
            coverTitle = wb.createCellStyle();
            XSSFFont cf = wb.createFont();
            cf.setBold(false); cf.setFontHeightInPoints((short) 22);
            coverTitle.setFont(cf);
            coverTitle.setAlignment(HorizontalAlignment.CENTER);
            coverTitle.setVerticalAlignment(VerticalAlignment.CENTER);

            sectionTitle = wb.createCellStyle();
            XSSFFont sf = wb.createFont();
            sf.setBold(true); sf.setFontHeightInPoints((short) 14);
            sectionTitle.setFont(sf);

            detailTitle = wb.createCellStyle();
            XSSFFont dtf = wb.createFont();
            dtf.setBold(false); dtf.setFontHeightInPoints((short) 14);
            detailTitle.setFont(dtf);
            detailTitle.setAlignment(HorizontalAlignment.CENTER);
            detailTitle.setVerticalAlignment(VerticalAlignment.CENTER);

            header = wb.createCellStyle();
            XSSFFont hf = wb.createFont();
            hf.setBold(true); hf.setFontHeightInPoints((short) 10);
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

            dataCenter = wb.createCellStyle();
            dataCenter.cloneStyleFrom(data);

            dataLeft = wb.createCellStyle();
            dataLeft.cloneStyleFrom(data);
            dataLeft.setAlignment(HorizontalAlignment.LEFT);

            approvalLabel = wb.createCellStyle();
            approvalLabel.setAlignment(HorizontalAlignment.CENTER);
            approvalLabel.setVerticalAlignment(VerticalAlignment.CENTER);
            approvalLabel.setWrapText(true);
            setBorder(approvalLabel);

            approvalHeader = wb.createCellStyle();
            approvalHeader.setAlignment(HorizontalAlignment.CENTER);
            approvalHeader.setVerticalAlignment(VerticalAlignment.CENTER);
            setBorder(approvalHeader);

            approvalBody = wb.createCellStyle();
            approvalBody.setAlignment(HorizontalAlignment.CENTER);
            approvalBody.setVerticalAlignment(VerticalAlignment.CENTER);
            setBorder(approvalBody);
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

    private void setRowHeight(XSSFSheet sheet, int rowIdx, float height) {
        getOrCreateRow(sheet, rowIdx).setHeightInPoints(height);
    }

    private double nvl(Double v) { return v != null ? v : 0.0; }

    /**
     * 병합 구간 전체 셀에 값 + 테두리 적용
     * POI는 병합 첫 셀에만 스타일이 적용되므로
     * 나머지 셀에도 직접 테두리를 칠해줘야 함
     */
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