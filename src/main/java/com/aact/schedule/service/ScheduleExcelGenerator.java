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
 * 제목: AACT {년}년 {월}월 {팀명} / {편명} {status} SKD
 *
 * 헤더 구조 (2행):
 * row 0: 순번 | 일자 | 1  | 2  | ... | 말일 | 휴무일수 | 사용휴무 | 사용연차 | 잔여연차
 * row 1:      | 성명 | 토 | 일 | ... | 요일 |          |          |          |
 *
 * 데이터:
 * 순번(자동) | 성명(DB) | 각 날짜별 값 | 집계 4개
 */
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

        YearMonth ym    = YearMonth.parse(applyYearMonth);
        int lastDay     = ym.lengthOfMonth();
        int year        = ym.getYear();
        int month       = ym.getMonthValue();

        try (XSSFWorkbook wb = new XSSFWorkbook();
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {

            XSSFSheet sheet = wb.createSheet(department + " " + month + "월 SKD");
            StyleSet styles = new StyleSet(wb);

            // ── row 0: 제목 + 결재란 헤더
            createTitleRow(sheet, styles, year, month, department, flightCode, status, approvers, lastDay);

            // ── row 1: 결재자 서명칸
            createApprovalRow(sheet, styles, approvers, lastDay);

            // ── row 2~3: 헤더 2행
            // row 2: 순번 | 일자 | 1 | 2 | ... | 말일 | 휴무일수 | 사용휴무 | 사용연차 | 잔여연차
            // row 3:      | 성명 | 토 | 일 | ... | 요일 |
            createHeaderRows(sheet, styles, lastDay, ym);

            // ── row 4~: 직원별 데이터 (순번 자동)
            int rowIdx = 4;
            int seq = 1;
            for (ScheduleRecord rec : records) {
                Row dr = sheet.createRow(rowIdx++);
                Map<Integer, String> days = scheduleRecordMapper.mapDaysFromEntity(rec);
                createDataRow(dr, styles, rec, days, lastDay, seq++);
            }

            // 열 너비
            sheet.setColumnWidth(0, 1500);  // 순번
            sheet.setColumnWidth(1, 3500);  // 일자/성명
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
                                String flightCode, String status,
                                List<String> approvers, int lastDay) {
        Row r0 = sheet.createRow(0);
        int totalCols    = 2 + lastDay + 4;
        int approvalCols = approvers != null ? approvers.size() * 2 : 0;
        int titleEnd     = totalCols - approvalCols - 1;

        // 제목 (status 동적 - "확정"/"임시" 등)
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, titleEnd));
        Cell titleCell = r0.createCell(0);
        titleCell.setCellValue(
                "AACT " + year + "년 " + month + "월 " + department
                        + " / " + flightCode + " " + status + " SKD");
        titleCell.setCellStyle(styles.title);

        // 결재란 헤더 (동적)
        if (approvers != null && !approvers.isEmpty()) {
            int col = titleEnd + 1;
            for (String approver : approvers) {
                sheet.addMergedRegion(new CellRangeAddress(0, 0, col, col + 1));
                Cell c = r0.createCell(col);
                c.setCellValue(approver);
                c.setCellStyle(styles.approvalHeader);
                col += 2;
            }
        }
    }

    // ── row 1: 결재자 서명칸
    private void createApprovalRow(XSSFSheet sheet, StyleSet styles,
                                   List<String> approvers, int lastDay) {
        Row r1 = sheet.createRow(1);
        r1.setHeightInPoints(45);
        if (approvers == null || approvers.isEmpty()) return;

        int totalCols = 2 + lastDay + 4;
        int col = totalCols - (approvers.size() * 2);
        for (String ignored : approvers) {
            sheet.addMergedRegion(new CellRangeAddress(1, 1, col, col + 1));
            r1.createCell(col).setCellStyle(styles.approvalBody);
            col += 2;
        }
    }

    // ── row 2~3: 헤더 2행
    // row 2: 순번 | 일자 | 1  | 2  | ... | 말일 | 휴무일수 | 사용휴무 | 사용연차 | 잔여연차
    // row 3:      | 성명 | 토 | 일 | ... | 요일 |
    private void createHeaderRows(XSSFSheet sheet, StyleSet styles,
                                  int lastDay, YearMonth ym) {
        Row r2 = sheet.createRow(2);
        Row r3 = sheet.createRow(3);
        r2.setHeightInPoints(20);
        r3.setHeightInPoints(20);

        String[] dayNames = {"일", "월", "화", "수", "목", "금", "토"};

        // 순번: row2~3 병합
        sheet.addMergedRegion(new CellRangeAddress(2, 3, 0, 0));
        setCell(r2, 0, "순번", styles.header);

        // 일자(row2) / 성명(row3)
        setCell(r2, 1, "일자", styles.header);
        setCell(r3, 1, "성명", styles.header);

        // 날짜(row2) / 요일(row3)
        for (int d = 1; d <= lastDay; d++) {
            LocalDate date = ym.atDay(d);
            String dow = dayNames[date.getDayOfWeek().getValue() % 7];

            // 주말 여부
            boolean isWeekend = date.getDayOfWeek().getValue() == 6  // 토
                    || date.getDayOfWeek().getValue() % 7 == 0;       // 일

            CellStyle dateStyle = isWeekend ? styles.headerWeekend : styles.header;
            CellStyle dowStyle  = isWeekend ? styles.headerWeekend : styles.header;

            setCell(r2, 1 + d, String.valueOf(d), dateStyle); // 날짜
            setCell(r3, 1 + d, dow, dowStyle);                // 요일
        }

        // 집계 컬럼 (row2~3 병합)
        int tail = 2 + lastDay;
        String[] tails = {"휴무\n일수", "사용\n휴무", "사용\n연차", "잔여\n연차"};
        for (int i = 0; i < 4; i++) {
            sheet.addMergedRegion(new CellRangeAddress(2, 3, tail + i, tail + i));
            setCell(r2, tail + i, tails[i], styles.header);
        }
    }

    // ── 데이터 행
    private void createDataRow(Row row, StyleSet styles,
                               ScheduleRecord rec, Map<Integer, String> days,
                               int lastDay, int seq) {
        // 순번 자동
        setCell(row, 0, String.valueOf(seq), styles.data);
        // 성명
        setCell(row, 1, rec.getEmployeeName(), styles.data);

        // 날짜별 값
        for (int d = 1; d <= lastDay; d++) {
            String val = days.getOrDefault(d, "");
            CellStyle style = "X".equals(val) ? styles.dayOff : styles.data;
            setCell(row, 1 + d, val, style);
        }

        // 집계
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
            tint(header, (byte)217, (byte)217, (byte)217);

            headerWeekend = base(wb, true, 10, true);
            tint(headerWeekend, (byte)255, (byte)230, (byte)230);

            data = base(wb, false, 10, false);

            dayOff = base(wb, false, 10, false);
            tint(dayOff, (byte)242, (byte)242, (byte)242);

            approvalHeader = base(wb, true, 10, false);
            ((XSSFCellStyle) approvalHeader).setAlignment(HorizontalAlignment.CENTER);

            approvalBody = base(wb, false, 10, false);
        }

        private CellStyle base(XSSFWorkbook wb, boolean bold, int size, boolean wrap) {
            XSSFCellStyle s = wb.createCellStyle();
            XSSFFont f = wb.createFont();
            f.setBold(bold); f.setFontHeightInPoints((short) size);
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