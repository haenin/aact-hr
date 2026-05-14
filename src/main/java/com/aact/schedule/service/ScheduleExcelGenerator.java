package com.aact.schedule.service;

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

/**
 * DB 데이터 → 월간 스케줄표 엑셀 생성
 *
 * 제목: AACT {년}년 {월}월 {팀명} / {편명} 확정 SKD
 * 결재: 동적 (부장/상무 등 개수 무관)
 * 컬럼: 순번 | 성명 | 1일~말일(요일) | 휴무일수 | 사용휴무 | 사용연차 | 잔여연차
 */
@Component
@RequiredArgsConstructor
public class ScheduleExcelGenerator {

    private final ScheduleRecordRepository scheduleRecordRepository;
    private final ScheduleRecordMapper scheduleRecordMapper;

    public byte[] generate(String applyYearMonth, String department,
                           String flightCode, List<String> approvers) throws Exception {

        List<ScheduleRecord> records = scheduleRecordRepository
                .findByApplyYearMonthAndDepartmentOrderBySeqAsc(applyYearMonth, department);

        YearMonth ym = YearMonth.parse(applyYearMonth);
        int lastDay = ym.lengthOfMonth(); // 3월=31, 4월=30, 2월=28or29 자동 계산
        int year    = ym.getYear();
        int month   = ym.getMonthValue();

        try (XSSFWorkbook wb = new XSSFWorkbook();
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {

            XSSFSheet sheet = wb.createSheet(department + " " + month + "월 SKD");
            StyleSet styles = new StyleSet(wb);

            // row 0: 제목 + 결재란 헤더
            createTitleRow(sheet, styles, year, month, department, flightCode, approvers, lastDay);

            // row 1: 결재자 서명칸 (도장 자리)
            createApprovalRow(sheet, styles, approvers, lastDay);

            // row 2: 헤더 (순번 | 성명 | 1일~말일+요일 | 집계 4개)
            createHeaderRow(sheet, styles, lastDay, ym);

            // row 3~: 직원별 데이터
            int rowIdx = 3;
            for (ScheduleRecord rec : records) {
                Row dr = sheet.createRow(rowIdx++);
                Map<Integer, String> days = scheduleRecordMapper.mapDaysFromEntity(rec);
                createDataRow(dr, styles, rec, days, lastDay);
            }

            // 열 너비
            sheet.setColumnWidth(0, 1500);  // 순번
            sheet.setColumnWidth(1, 3500);  // 성명
            for (int i = 2; i < 2 + lastDay; i++) {
                sheet.setColumnWidth(i, 1200);
            }
            int tail = 2 + lastDay;
            sheet.setColumnWidth(tail,     2200);
            sheet.setColumnWidth(tail + 1, 2200);
            sheet.setColumnWidth(tail + 2, 2200);
            sheet.setColumnWidth(tail + 3, 2200);

            wb.write(out);
            return out.toByteArray();
        }
    }

    // ── row 0: 제목 + 결재란 헤더
    private void createTitleRow(XSSFSheet sheet, StyleSet styles,
                                 int year, int month, String department,
                                 String flightCode, List<String> approvers, int lastDay) {
        Row r0 = sheet.createRow(0);
        int totalCols   = 2 + lastDay + 4;
        int approvalCols = approvers.size() * 2;
        int titleEnd    = totalCols - approvalCols - 1;

        // 제목
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, titleEnd));
        Cell titleCell = r0.createCell(0);
        titleCell.setCellValue(
                "AACT " + year + "년 " + month + "월 " + department + " / " + flightCode + " 확정 SKD");
        titleCell.setCellStyle(styles.title);

        // 결재란 헤더 (동적)
        int col = titleEnd + 1;
        for (String approver : approvers) {
            sheet.addMergedRegion(new CellRangeAddress(0, 0, col, col + 1));
            Cell c = r0.createCell(col);
            c.setCellValue(approver);
            c.setCellStyle(styles.approvalHeader);
            col += 2;
        }
    }

    // ── row 1: 결재자 서명칸
    private void createApprovalRow(XSSFSheet sheet, StyleSet styles,
                                    List<String> approvers, int lastDay) {
        Row r1 = sheet.createRow(1);
        int totalCols = 2 + lastDay + 4;
        int col = totalCols - (approvers.size() * 2);
        for (int i = 0; i < approvers.size(); i++) {
            sheet.addMergedRegion(new CellRangeAddress(1, 1, col, col + 1));
            r1.createCell(col).setCellStyle(styles.approvalBody);
            col += 2;
        }
        r1.setHeightInPoints(45);
    }

    // ── row 2: 헤더
    private void createHeaderRow(XSSFSheet sheet, StyleSet styles,
                                  int lastDay, YearMonth ym) {
        Row r2 = sheet.createRow(2);
        String[] dayNames = {"일", "월", "화", "수", "목", "금", "토"};

        setCell(r2, 0, "순번", styles.header);
        setCell(r2, 1, "성명", styles.header);

        for (int d = 1; d <= lastDay; d++) {
            LocalDate date = ym.atDay(d);
            String dow = dayNames[date.getDayOfWeek().getValue() % 7];
            // 토(6번인덱스)=파랑, 일(0번인덱스)=빨강 구분
            CellStyle dayStyle = (date.getDayOfWeek().getValue() == 6 || date.getDayOfWeek().getValue() % 7 == 0)
                    ? styles.headerWeekend : styles.header;
            setCell(r2, 1 + d, d + "\n" + dow, dayStyle);
        }

        int tail = 2 + lastDay;
        setCell(r2, tail,     "휴무\n일수", styles.header);
        setCell(r2, tail + 1, "사용\n휴무", styles.header);
        setCell(r2, tail + 2, "사용\n연차", styles.header);
        setCell(r2, tail + 3, "잔여\n연차", styles.header);

        r2.setHeightInPoints(30);
    }

    // ── 데이터 행
    private void createDataRow(Row row, StyleSet styles,
                                ScheduleRecord rec, Map<Integer, String> days, int lastDay) {
        setCell(row, 0, String.valueOf(rec.getSeq()), styles.data);
        setCell(row, 1, rec.getEmployeeName(), styles.data);

        for (int d = 1; d <= lastDay; d++) {
            String val = days.getOrDefault(d, "");
            CellStyle style = "X".equals(val) ? styles.dayOff : styles.data;
            setCell(row, 1 + d, val, style);
        }

        int tail = 2 + lastDay;
        setCell(row, tail,     rec.getOffDays()      != null ? String.valueOf(rec.getOffDays())      : "", styles.data);
        setCell(row, tail + 1, rec.getUsedOff()      != null ? String.valueOf(rec.getUsedOff())      : "", styles.data);
        setCell(row, tail + 2, rec.getUsedAnnual()   != null ? String.valueOf(rec.getUsedAnnual())   : "", styles.data);
        setCell(row, tail + 3, rec.getRemainAnnual() != null ? String.valueOf(rec.getRemainAnnual()) : "", styles.data);
    }

    // ── 스타일
    private static class StyleSet {
        final CellStyle title, header, headerWeekend, data, dayOff, approvalHeader, approvalBody;

        StyleSet(XSSFWorkbook wb) {
            title = base(wb, true, 12, false);
            ((XSSFCellStyle) title).setAlignment(HorizontalAlignment.CENTER);

            header = base(wb, true, 10, true);
            tint(wb, header, (byte)217, (byte)217, (byte)217);

            headerWeekend = base(wb, true, 10, true);
            tint(wb, headerWeekend, (byte)255, (byte)230, (byte)230);

            data = base(wb, false, 10, false);

            dayOff = base(wb, false, 10, false);
            tint(wb, dayOff, (byte)242, (byte)242, (byte)242);

            approvalHeader = base(wb, true, 10, false);
            ((XSSFCellStyle) approvalHeader).setAlignment(HorizontalAlignment.CENTER);

            approvalBody = base(wb, false, 10, false);
        }

        private CellStyle base(XSSFWorkbook wb, boolean bold, int size, boolean wrap) {
            XSSFCellStyle s = wb.createCellStyle();
            XSSFFont f = wb.createFont();
            f.setBold(bold);
            f.setFontHeightInPoints((short) size);
            s.setFont(f);
            s.setAlignment(HorizontalAlignment.CENTER);
            s.setVerticalAlignment(VerticalAlignment.CENTER);
            s.setWrapText(wrap);
            s.setBorderTop(BorderStyle.THIN);
            s.setBorderBottom(BorderStyle.THIN);
            s.setBorderLeft(BorderStyle.THIN);
            s.setBorderRight(BorderStyle.THIN);
            return s;
        }

        private void tint(XSSFWorkbook wb, CellStyle s, byte r, byte g, byte b) {
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
