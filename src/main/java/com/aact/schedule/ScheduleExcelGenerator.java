package com.aact.schedule;

import com.aact.schedule.entity.ScheduleRecord;
import com.aact.schedule.mapper.ScheduleRecordMapper;
import com.aact.schedule.repository.ScheduleRecordRepository;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Component;

import java.io.ByteArrayOutputStream;
import java.time.LocalDate;
import java.time.YearMonth;
import java.util.List;
import java.util.Map;

@Component
@RequiredArgsConstructor
public class ScheduleExcelGenerator {

    private final ScheduleRecordRepository scheduleRecordRepository;
    private final ScheduleRecordMapper scheduleRecordMapper;

    public byte[] generate(String applyYearMonth, String department,
                           String flightCode, String status,
                           List<String> approvers) throws Exception {

        List<ScheduleRecord> records = scheduleRecordRepository
                .findByApplyYearMonthAndDepartmentOrderBySeqAsc(applyYearMonth, department);

        YearMonth ym = YearMonth.parse(applyYearMonth);
        int lastDay = ym.lengthOfMonth();
        int year = ym.getYear();
        int month = ym.getMonthValue();

        try (XSSFWorkbook wb = new XSSFWorkbook();
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {

            String sheetName = department + " " + month + "월 SKD";
            sheetName = sheetName.replaceAll("[\\\\*?/\\[\\]]", "");
            if (sheetName.length() > 31) sheetName = sheetName.substring(0, 31);

            XSSFSheet sheet = wb.createSheet(sheetName);
            StyleSet styles = new StyleSet(wb);

            // 1. 제목 및 결재란 (0~1행 사용)
            createTitleAndApprovalHeader(sheet, styles, year, month, department, flightCode, status, approvers);

            // --- [수정] 결재란과 테이블 사이 간격 조절 ---
            Row emptyRow = sheet.createRow(2); // 2번 행 생성
            emptyRow.setHeightInPoints(40);    // 기본값은 보통 15~20pt입니다. 40~60 정도로 키우면 확 벌어집니다.
// ----------------------------------------

            // 2. 헤더 (3~4행 사용 / 2행은 빈칸)
            createHeaderRows(sheet, styles, lastDay, ym);

            // 3. 데이터 (5행부터 시작)
            int rowIdx = 5;
            int seq = 1;
            for (ScheduleRecord rec : records) {
                Row dr = sheet.createRow(rowIdx++);
                dr.setHeightInPoints(25); //
                Map<Integer, String> days = scheduleRecordMapper.mapDaysFromEntity(rec);
                createDataRow(dr, styles, rec, days, lastDay, seq++, ym);
            }


            applySheetSettings(sheet, lastDay);
            wb.write(out);
            return out.toByteArray();
        }
    }

    private void createTitleAndApprovalHeader(XSSFSheet sheet, StyleSet styles,
                                              int year, int month, String department,
                                              String flightCode, String status,
                                              List<String> approvers) {
        Row r0 = sheet.createRow(0);
        Row r1 = sheet.createRow(1);
        r0.setHeightInPoints(35);
        r1.setHeightInPoints(55);

        int approvalCount = (approvers != null) ? approvers.size() : 0;

        int approvalStartCol = 32; // AG (32)
        int approvalEndCol = 36;   // AK (36)
        int availableWidth = approvalEndCol - approvalStartCol;

        // --- [해결] 실제 병합은 AF(31)까지만 진행하여 결재란(32~)과 겹치지 않게 방어 ---
        sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, approvalStartCol - 1));
        Cell titleCell = r0.createCell(0);
        titleCell.setCellValue("                                           AACT " + year + "년 " + month + "월 " + department
                + " / " + flightCode + " " + status + " SKD");

        // 스타일을 통해 시각적으로만 AK까지의 중앙에 정렬되도록 트릭 적용
        titleCell.setCellStyle(styles.titleNoBorder);

        // --- 결재란 로직 (이제 제목과 절대 충돌하지 않음) ---
        if (approvalCount > 0) {
            sheet.addMergedRegion(new CellRangeAddress(0, 1, approvalStartCol, approvalStartCol));
            Cell lCell = r0.createCell(approvalStartCol);
            lCell.setCellValue("결\n재");
            lCell.setCellStyle(styles.approvalLabel);
            r1.createCell(approvalStartCol).setCellStyle(styles.approvalLabel);

            if (approvalCount == 1) {
                mergeAndStyleApprovalCell(r0, r1, sheet, styles, approvalStartCol + 1, approvalEndCol, approvers.get(0));
            } else if (approvalCount == 2) {
                int colFrom1 = approvalStartCol + 1;
                int colTo1 = colFrom1 + 1;
                mergeAndStyleApprovalCell(r0, r1, sheet, styles, colFrom1, colTo1, approvers.get(0));

                int colFrom2 = colTo1 + 1;
                int colTo2 = approvalEndCol;
                mergeAndStyleApprovalCell(r0, r1, sheet, styles, colFrom2, colTo2, approvers.get(1));
            } else if (approvalCount == 3) {
                int col1 = approvalStartCol + 1;
                mergeAndStyleApprovalCell(r0, r1, sheet, styles, col1, col1, approvers.get(0));

                int col2 = col1 + 1;
                mergeAndStyleApprovalCell(r0, r1, sheet, styles, col2, col2, approvers.get(1));

                int colFrom3 = col2 + 1;
                int colTo3 = approvalEndCol;
                mergeAndStyleApprovalCell(r0, r1, sheet, styles, colFrom3, colTo3, approvers.get(2));
            } else {
                for (int i = 0; i < Math.min(approvalCount, availableWidth); i++) {
                    int col = approvalStartCol + 1 + i;
                    mergeAndStyleApprovalCell(r0, r1, sheet, styles, col, col, approvers.get(i));
                }
            }
        }
    }

    private void mergeAndStyleApprovalCell(Row r0, Row r1, XSSFSheet sheet, StyleSet styles,
                                           int colFrom, int colTo, String approverName) {
        if (colFrom != colTo) {
            sheet.addMergedRegion(new CellRangeAddress(0, 0, colFrom, colTo));
            sheet.addMergedRegion(new CellRangeAddress(1, 1, colFrom, colTo));
        }

        for (int c = colFrom; c <= colTo; c++) {
            r0.createCell(c).setCellStyle(styles.approvalHeader);
            r1.createCell(c).setCellStyle(styles.approvalBody);
        }

        r0.getCell(colFrom).setCellValue(approverName);
    }

    private void createHeaderRows(XSSFSheet sheet, StyleSet styles, int lastDay, YearMonth ym) {
        Row r3 = sheet.createRow(3);
        Row r4 = sheet.createRow(4);
        r3.setHeightInPoints(25);
        r4.setHeightInPoints(25);
        String[] dayNames = {"일", "월", "화", "수", "목", "금", "토"};

        sheet.addMergedRegion(new CellRangeAddress(3, 4, 0, 0));
        setCell(r3, 0, "순번", styles.header);
        r4.createCell(0).setCellStyle(styles.header);

        setCell(r3, 1, "일자", styles.header);
        setCell(r4, 1, "성명", styles.header);

        for (int d = 1; d <= lastDay; d++) {
            LocalDate date = ym.atDay(d);
            String dow = dayNames[date.getDayOfWeek().getValue() % 7];
            setCell(r3, 1 + d, String.valueOf(d), styles.header);
            setCell(r4, 1 + d, dow, styles.header);
        }

        int tail = 2 + lastDay;
        String[] tails = {"휴무일수", "사용휴무", "사용연차", "잔여연차"};
        for (int i = 0; i < 4; i++) {
            int colIdx = tail + i;
            sheet.addMergedRegion(new CellRangeAddress(3, 4, colIdx, colIdx));
            setCell(r3, colIdx, tails[i], styles.header);
            r4.createCell(colIdx).setCellStyle(styles.header);
        }
    }

    private void createDataRow(Row row, StyleSet styles, ScheduleRecord rec, Map<Integer, String> days, int lastDay, int seq, YearMonth ym) {
        setCell(row, 0, String.valueOf(seq), styles.data);
        setCell(row, 1, rec.getEmployeeName(), styles.data);

        for (int d = 1; d <= lastDay; d++) {
            String val = days.getOrDefault(d, "");

            LocalDate date = ym.atDay(d);
            int dayOfWeek = date.getDayOfWeek().getValue();
            boolean isWeekend = (dayOfWeek == 6 || dayOfWeek == 7);

            setCell(row, 1 + d, val, styles.data);
        }

        int tail = 2 + lastDay;
        setCell(row, tail, String.valueOf(rec.getOffDays() != null ? rec.getOffDays() : ""), styles.data);
        setCell(row, tail + 1, String.valueOf(rec.getUsedOff() != null ? rec.getUsedOff() : ""), styles.data);
        setCell(row, tail + 2, String.valueOf(rec.getUsedAnnual() != null ? rec.getUsedAnnual() : ""), styles.data);
        setCell(row, tail + 3, String.valueOf(rec.getRemainAnnual() != null ? rec.getRemainAnnual() : ""), styles.data);
    }

    private void applySheetSettings(XSSFSheet sheet, int lastDay) {
        // [수정] 순번 열 너비도 1200으로 조정 (공간 확보)
        sheet.setColumnWidth(0, 1200);
        sheet.setColumnWidth(1, 3500);
        for (int i = 2; i < 2 + lastDay; i++) sheet.setColumnWidth(i, 1200);

        int tail = 2 + lastDay;
        for (int i = 0; i < 4; i++) sheet.setColumnWidth(tail + i, 2200);

        sheet.setDisplayGridlines(false);

        PrintSetup ps = sheet.getPrintSetup();
        ps.setPaperSize(PrintSetup.A4_PAPERSIZE);
        ps.setLandscape(true); // 가로 인쇄

        // --- [추가] 인쇄 시 모든 열을 한 페이지에 맞춤 설정 ---
        sheet.setFitToPage(true);
        ps.setFitWidth((short) 1);  // 가로(너비)를 1페이지에 맞춤
        ps.setFitHeight((short) 0); // 세로는 데이터 양에 따라 여러 장으로 나옴
        // --------------------------------------------------

        // 인쇄 영역 설정 (A열부터 마지막 집계열까지)
        sheet.getWorkbook().setPrintArea(0, 0, tail + 3, 0, sheet.getLastRowNum());

        sheet.setMargin(Sheet.LeftMargin, 0.3);
        sheet.setMargin(Sheet.RightMargin, 0.3);
        sheet.setMargin(Sheet.TopMargin, 0.5);
        sheet.setMargin(Sheet.BottomMargin, 0.5);
    }
    private static class StyleSet {
        final CellStyle titleNoBorder, header, data, dayOff, noBorder, approvalLabel, approvalHeader, approvalBody;

        StyleSet(XSSFWorkbook wb) {
            noBorder = base(wb, false, 10, false, false);
            titleNoBorder = base(wb, true, 26, false, false);
            // [수정] 일반 CENTER 대신 CENTER_SELECTION을 주어 우측 결재란 전체를 포괄하는 가상 중앙정렬 구현
            ((XSSFCellStyle) titleNoBorder).setAlignment(HorizontalAlignment.CENTER_SELECTION);

            // 헤더 폰트 크기: 11pt 고정
            header = base(wb, true, 11, false, true);
            tint(header, (byte)217, (byte)217, (byte)217);

            data = base(wb, false, 11, false, true);
            dayOff = base(wb, false, 11, false, true);
            tint(dayOff, (byte)242, (byte)242, (byte)242);

            approvalLabel = base(wb, true, 10, true, true);
            tint(approvalLabel, (byte)217, (byte)217, (byte)217);

            approvalHeader = base(wb, true, 10, false, true);
            approvalBody = base(wb, false, 10, false, true);
        }

        private CellStyle base(XSSFWorkbook wb, boolean bold, int size, boolean wrap, boolean border) {
            XSSFCellStyle s = wb.createCellStyle();
            XSSFFont f = wb.createFont();
            f.setBold(bold); f.setFontHeightInPoints((short) size);
            s.setFont(f);
            s.setAlignment(HorizontalAlignment.CENTER);
            s.setVerticalAlignment(VerticalAlignment.CENTER);
            s.setWrapText(wrap);
            if (border) {
                s.setBorderTop(BorderStyle.THIN); s.setBorderBottom(BorderStyle.THIN);
                s.setBorderLeft(BorderStyle.THIN); s.setBorderRight(BorderStyle.THIN);
            }
            return s;
        }

        private void tint(CellStyle s, byte r, byte g, byte b) {
            ((XSSFCellStyle) s).setFillForegroundColor(new XSSFColor(new byte[]{r, g, b}, null));
            s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
    }

    private void setCell(Row row, int col, String value, CellStyle style) {
        Cell cell = row.createCell(col);
        cell.setCellValue(value != null ? value : "");
        cell.setCellStyle(style);
    }
}