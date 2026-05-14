package com.aact.overtime.service;

import com.aact.overtime.dto.ApplicationDto;
import com.aact.overtime.entity.OvertimePayApplication;
import com.aact.overtime.repository.ApplicationRepository;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Component;

import java.io.ByteArrayOutputStream;
import java.util.List;

/**
 * DB 데이터 → 시간외근무수당 신청서 엑셀 생성
 *
 * 구조:
 *  - 제목: 시간 외 근무수당 신청서({부서명})
 *  - 결재란: 동적 (담당/대리 등)
 *  - AACT 로고 + 년도/월
 *  - 테이블: NO / 일자 / 성명 / 예정근무시간 / 실근무시간 / 연장/야간/휴일 / 근무사유
 *  - 합계 행
 *  - 서명란: 신청자/팀장/부서장 등 동적
 */
@Component
@RequiredArgsConstructor
public class ApplicationExcelGenerator {

    private final ApplicationRepository overtimeApplicationRepository;

    private static final int MAX_ROWS = 28; // 테이블 최대 행 수

    public byte[] generate(ApplicationDto.ExcelRequest request) throws Exception {
        String yearMonth = request.getApplyYearMonth();
        String dept      = request.getDepartment();
        int year  = Integer.parseInt(yearMonth.substring(0, 4));
        int month = Integer.parseInt(yearMonth.substring(5));

        // DB 조회
        List<OvertimePayApplication> records = overtimeApplicationRepository
                .findByApplyYearMonthAndDepartmentOrderByNoAsc(yearMonth, dept);

        // 합계 조회
        List<Object[]> sums = overtimeApplicationRepository
                .sumByYearMonthAndDepartment(yearMonth, dept);
        double totalExt   = 0, totalNight = 0, totalHoliday = 0;
        if (!sums.isEmpty() && sums.get(0)[0] != null) {
            totalExt     = ((Number) sums.get(0)[0]).doubleValue();
            totalNight   = ((Number) sums.get(0)[1]).doubleValue();
            totalHoliday = ((Number) sums.get(0)[2]).doubleValue();
        }

        try (XSSFWorkbook wb = new XSSFWorkbook();
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {

            XSSFSheet sheet = wb.createSheet(month + "월 신청서");
            StyleSet styles = new StyleSet(wb);

            int rowIdx = 0;

            // ── row 0: 제목
            rowIdx = createTitleRow(sheet, styles, dept, request.getApprovers(), rowIdx);

            // ── row 1~2: AACT + 년도/월
            rowIdx = createLogoRow(sheet, styles, year, month, rowIdx);

            // ── row 3: 테이블 헤더
            rowIdx = createTableHeader(sheet, styles, rowIdx);

            // ── row 4~: 데이터 행 (최대 28행)
            int dataStartIdx = rowIdx;
            for (OvertimePayApplication rec : records) {
                if (rowIdx - dataStartIdx >= MAX_ROWS) break;
                createDataRow(sheet, styles, rec, rowIdx++);
            }

            // 빈 행 채우기 (28행 맞추기)
            while (rowIdx - dataStartIdx < MAX_ROWS) {
                createEmptyRow(sheet, styles, rowIdx++);
            }

            // ── 합계 행
            createTotalRow(sheet, styles, totalExt, totalNight, totalHoliday, rowIdx++);

            // ── 서명란 (신청자/팀장/부서장 동적)
            createSignerRows(sheet, styles, request.getSigners(), rowIdx);

            // 열 너비
            sheet.setColumnWidth(0, 1500);   // NO
            sheet.setColumnWidth(1, 3000);   // 일자
            sheet.setColumnWidth(2, 3000);   // 성명
            sheet.setColumnWidth(3, 4000);   // 예정근무시간
            sheet.setColumnWidth(4, 4500);   // 실근무시간
            sheet.setColumnWidth(5, 2500);   // 연장
            sheet.setColumnWidth(6, 2500);   // 야간
            sheet.setColumnWidth(7, 2500);   // 휴일
            sheet.setColumnWidth(8, 10000);  // 근무사유

            wb.write(out);
            return out.toByteArray();
        }
    }

    // ── 제목 + 결재란
    private int createTitleRow(XSSFSheet sheet, StyleSet styles,
                                String dept, List<String> approvers, int rowIdx) {
        Row r = sheet.createRow(rowIdx);
        r.setHeightInPoints(30);

        // 제목 (결재란 시작 전까지 병합)
        int titleEnd = approvers == null || approvers.isEmpty() ? 8 : 8 - (approvers.size() * 2);
        sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx, 0, titleEnd));
        Cell titleCell = r.createCell(0);
        titleCell.setCellValue("시간 외 근무수당 신청서(" + dept + ")");
        titleCell.setCellStyle(styles.title);

        // 결재란 헤더 (동적)
        if (approvers != null) {
            int col = titleEnd + 1;
            for (String approver : approvers) {
                sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx, col, col + 1));
                Cell c = r.createCell(col);
                c.setCellValue(approver);
                c.setCellStyle(styles.approvalHeader);
                col += 2;
            }
        }
        return ++rowIdx;
    }

    // ── AACT 로고 + 결재 서명칸 + 년도/월
    private int createLogoRow(XSSFSheet sheet, StyleSet styles,
                               int year, int month, int rowIdx) {
        // 결재 서명칸
        Row r1 = sheet.createRow(rowIdx);
        r1.setHeightInPoints(35);
        sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx + 1, 0, 2));
        Cell logoCell = r1.createCell(0);
        logoCell.setCellValue("AACT");
        logoCell.setCellStyle(styles.logo);
        rowIdx++;

        // 년도/월
        Row r2 = sheet.createRow(rowIdx);
        sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx, 0, 2));
        Cell ymCell = r2.createCell(0);
        ymCell.setCellValue(year + " 년 " + month + "월");
        ymCell.setCellStyle(styles.yearMonth);
        return ++rowIdx;
    }

    // ── 테이블 헤더
    private int createTableHeader(XSSFSheet sheet, StyleSet styles, int rowIdx) {
        Row r = sheet.createRow(rowIdx);
        r.setHeightInPoints(20);
        setCell(r, 0, "NO",       styles.header);
        setCell(r, 1, "일 자",    styles.header);
        setCell(r, 2, "성 명",    styles.header);
        setCell(r, 3, "예정근무시간", styles.header);
        setCell(r, 4, "실 근무시간", styles.header);
        setCell(r, 5, "연장근무\n시간", styles.header);
        setCell(r, 6, "야간근무\n시간", styles.header);
        setCell(r, 7, "휴일근무\n시간", styles.header);
        setCell(r, 8, "근무사유", styles.header);
        return ++rowIdx;
    }

    // ── 데이터 행
    private void createDataRow(XSSFSheet sheet, StyleSet styles,
                                OvertimePayApplication rec, int rowIdx) {
        Row r = sheet.createRow(rowIdx);
        setCell(r, 0, String.valueOf(rec.getNo()), styles.data);
        setCell(r, 1, rec.getWorkDate()      != null ? rec.getWorkDate()      : "", styles.data);
        setCell(r, 2, rec.getEmployeeName()  != null ? rec.getEmployeeName()  : "", styles.data);
        setCell(r, 3, rec.getScheduledTime() != null ? rec.getScheduledTime() : "", styles.data);
        setCell(r, 4, rec.getActualTime()    != null ? rec.getActualTime()    : "", styles.data);
        setNumCell(r, 5, rec.getExtensionHours(), styles.data);
        setNumCell(r, 6, rec.getNightHours(),     styles.data);
        setNumCell(r, 7, rec.getHolidayHours(),   styles.data);
        setCell(r, 8, rec.getWorkReason()    != null ? rec.getWorkReason()    : "", styles.dataLeft);
    }

    // ── 빈 행 (~ 표시)
    private void createEmptyRow(XSSFSheet sheet, StyleSet styles, int rowIdx) {
        Row r = sheet.createRow(rowIdx);
        setCell(r, 0, "", styles.data);
        setCell(r, 1, "", styles.data);
        setCell(r, 2, "", styles.data);
        setCell(r, 3, "~", styles.data);
        setCell(r, 4, "~", styles.data);
        setCell(r, 5, "", styles.data);
        setCell(r, 6, "", styles.data);
        setCell(r, 7, "", styles.data);
        setCell(r, 8, "", styles.data);
    }

    // ── 합계 행
    private void createTotalRow(XSSFSheet sheet, StyleSet styles,
                                 double ext, double night, double holiday, int rowIdx) {
        Row r = sheet.createRow(rowIdx);
        sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx, 0, 4));
        Cell totalLabel = r.createCell(0);
        totalLabel.setCellValue("합   계");
        totalLabel.setCellStyle(styles.header);
        setNumCell(r, 5, ext,     styles.header);
        setNumCell(r, 6, night,   styles.header);
        setNumCell(r, 7, holiday, styles.header);
        setCell(r, 8, "", styles.data);
    }

    // ── 서명란 (동적)
    private void createSignerRows(XSSFSheet sheet, StyleSet styles,
                                  List<ApplicationDto.Signer> signers, int rowIdx) {
        if (signers == null || signers.isEmpty()) return;

        rowIdx += 2; // 여백
        for (ApplicationDto.Signer signer : signers) {
            Row r = sheet.createRow(rowIdx++);
            r.setHeightInPoints(25);
            sheet.addMergedRegion(new CellRangeAddress(rowIdx - 1, rowIdx - 1, 5, 6));
            sheet.addMergedRegion(new CellRangeAddress(rowIdx - 1, rowIdx - 1, 7, 8));

            setCell(r, 5, signer.getRole() + " :", styles.signerRole);
            setCell(r, 7, signer.getName(),        styles.signerName);
        }
    }

    // ── 스타일
    private static class StyleSet {
        final CellStyle title, header, data, dataLeft;
        final CellStyle approvalHeader, logo, yearMonth;
        final CellStyle signerRole, signerName;

        StyleSet(XSSFWorkbook wb) {
            title = base(wb, true, 14);
            ((XSSFCellStyle) title).setAlignment(HorizontalAlignment.CENTER);

            header = base(wb, true, 10);
            header.setWrapText(true);
            tint(header, (byte)217, (byte)217, (byte)217);

            data = base(wb, false, 10);
            ((XSSFCellStyle) data).setAlignment(HorizontalAlignment.CENTER);

            dataLeft = base(wb, false, 10);
            ((XSSFCellStyle) dataLeft).setAlignment(HorizontalAlignment.LEFT);

            approvalHeader = base(wb, true, 10);
            ((XSSFCellStyle) approvalHeader).setAlignment(HorizontalAlignment.CENTER);

            logo = base(wb, true, 16);
            ((XSSFCellStyle) logo).setAlignment(HorizontalAlignment.LEFT);

            yearMonth = base(wb, false, 11);
            ((XSSFCellStyle) yearMonth).setAlignment(HorizontalAlignment.LEFT);

            signerRole = base(wb, false, 11);
            ((XSSFCellStyle) signerRole).setAlignment(HorizontalAlignment.RIGHT);
            signerRole.setBorderBottom(BorderStyle.NONE);
            signerRole.setBorderTop(BorderStyle.NONE);
            signerRole.setBorderLeft(BorderStyle.NONE);
            signerRole.setBorderRight(BorderStyle.NONE);

            signerName = base(wb, false, 11);
            ((XSSFCellStyle) signerName).setAlignment(HorizontalAlignment.LEFT);
            signerName.setBorderBottom(BorderStyle.NONE);
            signerName.setBorderTop(BorderStyle.NONE);
            signerName.setBorderLeft(BorderStyle.NONE);
            signerName.setBorderRight(BorderStyle.NONE);
        }

        private CellStyle base(XSSFWorkbook wb, boolean bold, int size) {
            XSSFCellStyle s = wb.createCellStyle();
            XSSFFont f = wb.createFont();
            f.setBold(bold);
            f.setFontHeightInPoints((short) size);
            s.setFont(f);
            s.setVerticalAlignment(VerticalAlignment.CENTER);
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

    private void setNumCell(Row row, int col, Double value, CellStyle style) {
        Cell cell = row.createCell(col);
        cell.setCellValue(value != null ? value : 0.0);
        cell.setCellStyle(style);
    }
}
