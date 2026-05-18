package com.aact.overtime.service;

import com.aact.overtime.dto.ApplicationDto;
import com.aact.overtime.entity.OvertimePayApplication;
import com.aact.overtime.repository.ApplicationRepository;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextBody;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextParagraph;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextParagraphProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.STTextAlignType;
import org.springframework.stereotype.Component;

import java.io.ByteArrayOutputStream;
import java.util.List;

@Component
@RequiredArgsConstructor
public class ApplicationExcelGenerator {

    private final ApplicationRepository overtimeApplicationRepository;

    public byte[] generate(ApplicationDto.ExcelRequest request) throws Exception {
        String yearMonth = request.getApplyYearMonth();
        String dept      = request.getDepartment();
        int year  = Integer.parseInt(yearMonth.substring(0, 4));
        int month = Integer.parseInt(yearMonth.substring(5));

        List<OvertimePayApplication> records = overtimeApplicationRepository
                .findByApplyYearMonthAndDepartmentOrderByNoAsc(yearMonth, dept);

        List<Object[]> sums = overtimeApplicationRepository
                .sumByYearMonthAndDepartment(yearMonth, dept);
        double totalExt = 0, totalNight = 0, totalHoliday = 0;
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

            // ── [row 0] 제목 (col 0~8)
            rowIdx = createTitleRow(sheet, styles, dept, rowIdx);
            // rowIdx == 1

            // ── [row 1~2] 로고(col 0~6) + 결재란 Drawing(col 8에만, row 1~2)
            rowIdx = createLogoAndApprovalBlock(sheet, styles, request.getApprovers(), rowIdx, wb);
            // rowIdx == 3

            // ── [row 3] 년도/월
            rowIdx = createYearMonthRow(sheet, styles, year, month, rowIdx);

            // ── [row 4] 테이블 헤더
            rowIdx = createTableHeader(sheet, styles, rowIdx);

            // ── [row 5~] 데이터 행
            for (OvertimePayApplication rec : records) {
                createDataRow(sheet, styles, rec, rowIdx++);
            }

            // ── 합계 행
            createTotalRow(sheet, styles, totalExt, totalNight, totalHoliday, rowIdx++);

            // ── 서명란 (col 8 아래, 합계 바로 밑)
            createSignerRows(sheet, styles, request.getSigners(), rowIdx);

            // ── 열 너비
            sheet.setColumnWidth(0, 1200);
            sheet.setColumnWidth(1, 2800);
            sheet.setColumnWidth(2, 2500);
            sheet.setColumnWidth(3, 3800);
            sheet.setColumnWidth(4, 3800);
            sheet.setColumnWidth(5, 2800);
            sheet.setColumnWidth(6, 2800);
            sheet.setColumnWidth(7, 2800);
            sheet.setColumnWidth(8, 10000);

            // ── 인쇄 설정
            wb.setPrintArea(0, 0, 8, 0, rowIdx + 5);
            sheet.setFitToPage(true);
            sheet.setAutobreaks(true);

            PrintSetup ps = sheet.getPrintSetup();
            ps.setPaperSize(PrintSetup.A4_PAPERSIZE);
            ps.setFitWidth((short) 1);
            ps.setFitHeight((short) 0);
            ps.setLandscape(false);

            sheet.setMargin(Sheet.LeftMargin,   0.2);
            sheet.setMargin(Sheet.RightMargin,  0.2);
            sheet.setMargin(Sheet.TopMargin,    0.3);
            sheet.setMargin(Sheet.BottomMargin, 0.3);
            sheet.setMargin(Sheet.HeaderMargin, 0.0);
            sheet.setMargin(Sheet.FooterMargin, 0.0);

            sheet.setDisplayGridlines(false);
            wb.write(out);
            return out.toByteArray();
        }
    }

    // ── [row 0] 제목
    private int createTitleRow(XSSFSheet sheet, StyleSet styles,
                               String dept, int rowIdx) {
        Row r = sheet.createRow(rowIdx);
        r.setHeightInPoints(80);

        for (int i = 0; i <= 8; i++) {
            Cell cell = r.createCell(i);
            cell.setCellStyle(styles.noBorder);
        }

        sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx, 0, 8));

        Cell c = r.getCell(0);
        c.setCellValue("시간 외 근무수당 신청서 (" + dept + ")");
        c.setCellStyle(styles.titleNoBorder);

        return ++rowIdx;
    }

    private int createLogoAndApprovalBlock(XSSFSheet sheet, StyleSet styles,
                                           List<String> approvers, int rowIdx,
                                           XSSFWorkbook wb) {
        Row row1 = sheet.createRow(rowIdx);
        row1.setHeightInPoints(40);
        sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx, 0, 6));
        for (int col = 0; col <= 6; col++) {
            row1.createCell(col).setCellStyle(styles.noBorder);
        }

        XSSFDrawing drawing = sheet.createDrawingPatriarch();

        try {
            var is = getClass().getResourceAsStream("/img.png");
            if (is != null) {
                byte[] bytes = is.readAllBytes();
                int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
                is.close();

                XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, rowIdx + 1, 3, rowIdx + 2);
                anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);
                drawing.createPicture(anchor, pictureIdx);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        row1.createCell(7).setCellStyle(styles.noBorder);
        row1.createCell(8).setCellStyle(styles.noBorder);

        Row row2 = sheet.createRow(rowIdx + 1);
        row2.setHeightInPoints(40);
        sheet.addMergedRegion(new CellRangeAddress(rowIdx + 1, rowIdx + 1, 0, 6));
        for (int col = 0; col <= 6; col++) {
            row2.createCell(col).setCellStyle(styles.noBorder);
        }
        row2.createCell(7).setCellStyle(styles.noBorder);
        row2.createCell(8).setCellStyle(styles.noBorder);

        List<String> headers = (approvers != null && !approvers.isEmpty())
                ? approvers : List.of("담당", "대리");

        int totalEmu  = 3_000_000;
        int offsetEmu = 0;
        int labelEmu  = 400_000;

        int count = headers.size();
        int eachEmu = (totalEmu - labelEmu) / count;

        XSSFClientAnchor labelAnchor = drawing.createAnchor(
                offsetEmu, 0,
                offsetEmu + labelEmu, 0,
                8, rowIdx,
                8, rowIdx + 2
        );
        XSSFTextBox labelBox = drawing.createTextbox(labelAnchor);
        applyBoxStyle(labelBox, wb, "결\n재", true, true);

        for (int i = 0; i < count; i++) {
            int currentStartX = offsetEmu + labelEmu + (i * eachEmu);
            int currentEndX   = currentStartX + eachEmu;

            XSSFClientAnchor hAnchor = drawing.createAnchor(
                    currentStartX, 0, currentEndX, 0,
                    8, rowIdx,
                    8, rowIdx + 1
            );
            XSSFTextBox hBox = drawing.createTextbox(hAnchor);
            applyBoxStyle(hBox, wb, headers.get(i), true, true);

            XSSFClientAnchor sAnchor = drawing.createAnchor(
                    currentStartX, 0, currentEndX, 0,
                    8, rowIdx + 1,
                    8, rowIdx + 2
            );
            XSSFTextBox sBox = drawing.createTextbox(sAnchor);
            applyBoxStyle(sBox, wb, "", false, false);
        }

        return rowIdx + 2;
    }

    /**
     * 텍스트박스 스타일 공통 헬퍼 — 맑은 고딕 적용
     */
    private void applyBoxStyle(XSSFTextBox box, XSSFWorkbook wb, String text, boolean isBold, boolean hasFill) {
        box.setLineStyleColor(0, 0, 0);
        box.setLineWidth(0.5);

        if (hasFill) {
            box.setFillColor(217, 217, 217);
        } else {
            box.setNoFill(true);
        }

        XSSFFont font = wb.createFont();
        font.setBold(isBold);
        font.setFontHeightInPoints((short) 9);
        font.setFontName("맑은 고딕"); // ← 추가

        XSSFRichTextString richText = new XSSFRichTextString(text);
        richText.applyFont(font);

        box.setText(richText);
        box.setVerticalAlignment(VerticalAlignment.CENTER);
        setCTTextCenter(box);
    }

    // ── 년도/월
    private int createYearMonthRow(XSSFSheet sheet, StyleSet styles,
                                   int year, int month, int rowIdx) {
        Row r = sheet.createRow(rowIdx);
        r.setHeightInPoints(32);
        sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx, 0, 8));
        Cell c = r.createCell(0);
        c.setCellValue(year + "년  " + String.format("%02d", month) + "월");
        c.setCellStyle(styles.yearMonth);
        return ++rowIdx;
    }

    // ── 테이블 헤더
    private int createTableHeader(XSSFSheet sheet, StyleSet styles, int rowIdx) {
        Row r = sheet.createRow(rowIdx);
        r.setHeightInPoints(25);
        setCell(r, 0, "NO",          styles.header);
        setCell(r, 1, "일 자",        styles.header);
        setCell(r, 2, "성 명",        styles.header);
        setCell(r, 3, "예정근무시간",   styles.header);
        setCell(r, 4, "실근무시간",    styles.header);
        setCell(r, 5, "연장근무시간",  styles.header);
        setCell(r, 6, "야간근무시간",  styles.header);
        setCell(r, 7, "휴일근무시간",  styles.header);
        setCell(r, 8, "근무사유",      styles.header);
        return ++rowIdx;
    }

    // ── 데이터 행
    private void createDataRow(XSSFSheet sheet, StyleSet styles,
                               OvertimePayApplication rec, int rowIdx) {
        Row r = sheet.createRow(rowIdx);
        r.setHeightInPoints(25);

        String originalDate = rec.getWorkDate();
        String formattedDate = "";

        if (originalDate != null && originalDate.contains("-")) {
            try {
                String[] parts = originalDate.split("-");
                int month = Integer.parseInt(parts[1]);
                int day   = Integer.parseInt(parts[2]);
                formattedDate = String.format("%02d월 %02d일", month, day);
            } catch (Exception e) {
                formattedDate = originalDate;
            }
        } else {
            formattedDate = (originalDate != null) ? originalDate : "";
        }

        setCell(r, 0, String.valueOf(rec.getNo()), styles.dataCenter);
        setCell(r, 1, formattedDate, styles.dataCenter);

        setCell   (r, 2, rec.getEmployeeName()  != null ? rec.getEmployeeName()  : "",  styles.dataCenter);
        setCell   (r, 3, rec.getScheduledTime() != null ? rec.getScheduledTime() : "",  styles.dataCenter);
        setCell   (r, 4, rec.getActualTime()    != null ? rec.getActualTime()    : "",  styles.dataCenter);
        setNumCell(r, 5, rec.getExtensionHours(),  styles.dataCenter);
        setNumCell(r, 6, rec.getNightHours(),      styles.dataCenter);
        setNumCell(r, 7, rec.getHolidayHours(),    styles.dataCenter);
        setCell   (r, 8, rec.getWorkReason()    != null ? rec.getWorkReason()    : "",  styles.dataCenter);
    }

    // ── 합계 행
    private void createTotalRow(XSSFSheet sheet, StyleSet styles,
                                double ext, double night, double holiday, int rowIdx) {
        Row r = sheet.createRow(rowIdx);
        r.setHeightInPoints(22);

        for (int i = 0; i <= 8; i++) {
            Cell c = r.createCell(i);
            if (i <= 4) {
                c.setCellStyle(styles.totalLabel);
            } else if (i >= 5 && i <= 7) {
                c.setCellStyle(styles.totalValue);
            } else {
                c.setCellStyle(styles.header);
            }
        }

        sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx, 0, 4));

        r.getCell(0).setCellValue("합    계");
        r.getCell(5).setCellValue(ext);
        r.getCell(6).setCellValue(night);
        r.getCell(7).setCellValue(holiday);
    }

    /**
     * 서명란 — 맑은 고딕 적용
     */
    private void createSignerRows(XSSFSheet sheet, StyleSet styles,
                                  List<ApplicationDto.Signer> signers, int rowIdx) {
        if (signers == null || signers.isEmpty()) return;

        int startRow = rowIdx + 1;
        Workbook wb = sheet.getWorkbook();

        // ← 맑은 고딕 적용
        XSSFFont boldFont = (XSSFFont) wb.createFont();
        boldFont.setBold(true);
        boldFont.setFontHeightInPoints((short) 11);
        boldFont.setFontName("맑은 고딕"); // ← 추가

        XSSFFont plainFont = (XSSFFont) wb.createFont();
        plainFont.setBold(false);
        plainFont.setFontHeightInPoints((short) 11);
        plainFont.setFontName("맑은 고딕"); // ← 추가

        for (int i = 0; i < signers.size(); i++) {
            ApplicationDto.Signer signer = signers.get(i);
            Row r = sheet.createRow(startRow + i);
            r.setHeightInPoints(36);

            for (int col = 0; col <= 6; col++) {
                r.createCell(col).setCellStyle(styles.noBorder);
            }

            Cell roleCell = r.createCell(7);
            String rawRole = signer.getRole() != null ? signer.getRole() : "";

            String displayRole = String.join(" ", rawRole.split(""));
            if (rawRole.length() == 2) {
                displayRole = rawRole.charAt(0) + "    " + rawRole.charAt(1);
            }

            roleCell.setCellValue(displayRole + " :");
            roleCell.setCellStyle(styles.signerRole);

            Cell nameCell = r.createCell(8);
            XSSFRichTextString richText = new XSSFRichTextString();
            richText.append("  " + (signer.getName() != null ? signer.getName() : ""), boldFont);
            richText.append("                                            (인)", plainFont);

            nameCell.setCellValue(richText);
            nameCell.setCellStyle(styles.signerLine);
        }
    }

    // ────────────────────────────────────────────────────────────────
    // StyleSet — bordered(), plain() 모두 맑은 고딕 적용
    // ────────────────────────────────────────────────────────────────
    private static class StyleSet {
        final CellStyle title, titleNoBorder, approvalLabel, approvalHeader, approvalSign;
        final CellStyle logo, yearMonth, header, totalLabel;
        final CellStyle totalValue;
        final CellStyle dataCenter, dataLeft, signerLine, noBorder;
        final CellStyle signerRole, signerName, signerIn;

        StyleSet(XSSFWorkbook wb) {
            title = bordered(wb, true, 38);
            align(title, HorizontalAlignment.CENTER);

            approvalLabel = bordered(wb, true, 11);
            align(approvalLabel, HorizontalAlignment.CENTER);
            tint(approvalLabel, (byte)217, (byte)217, (byte)217);

            approvalHeader = bordered(wb, true, 9);
            align(approvalHeader, HorizontalAlignment.CENTER);
            tint(approvalHeader, (byte)217, (byte)217, (byte)217);

            approvalSign = bordered(wb, false, 9);
            align(approvalSign, HorizontalAlignment.CENTER);

            logo = plain(wb, true, 15);
            align(logo, HorizontalAlignment.LEFT);

            yearMonth = plain(wb, true, 13);
            align(yearMonth, HorizontalAlignment.LEFT);

            header = bordered(wb, true, 9);
            align(header, HorizontalAlignment.CENTER);
            header.setWrapText(false);
            tint(header, (byte)217, (byte)217, (byte)217);

            totalLabel = bordered(wb, true, 11);
            align(totalLabel, HorizontalAlignment.CENTER);
            tint(totalLabel, (byte)230, (byte)230, (byte)230);

            totalValue = bordered(wb, true, 11);
            align(totalValue, HorizontalAlignment.CENTER);
            tint(totalValue, (byte)217, (byte)217, (byte)217);

            dataCenter = bordered(wb, false, 11);
            align(dataCenter, HorizontalAlignment.CENTER);
            // 데이터 셀 폰트를 스케줄 인쇄 비율 기준으로 맞춤 (돋움 10pt)
            XSSFFont dataFont = wb.createFont();
            dataFont.setBold(false);
            dataFont.setFontHeightInPoints((short) 10);
            dataFont.setFontName("돋움");
            ((XSSFCellStyle) dataCenter).setFont(dataFont);

            CellStyle dateStyle = wb.createCellStyle();
            dateStyle.cloneStyleFrom(dataCenter);
            dateStyle.setDataFormat(wb.createDataFormat().getFormat("m\"월\" d\"일\""));

            dataLeft = bordered(wb, false, 10);
            align(dataLeft, HorizontalAlignment.LEFT);

            signerLine = plain(wb, false, 11);
            align(signerLine, HorizontalAlignment.LEFT);

            signerRole = plain(wb, false, 11);
            align(signerRole, HorizontalAlignment.RIGHT);

            signerName = plain(wb, false, 11);
            align(signerName, HorizontalAlignment.LEFT);

            signerIn = plain(wb, false, 11);
            align(signerIn, HorizontalAlignment.CENTER);

            noBorder = plain(wb, false, 10);

            titleNoBorder = plain(wb, true, 24);
            align(titleNoBorder, HorizontalAlignment.CENTER);
        }

        // ← setFontName("맑은 고딕") 추가
        private static CellStyle bordered(XSSFWorkbook wb, boolean bold, int size) {
            XSSFCellStyle s = wb.createCellStyle();
            XSSFFont f = wb.createFont();
            f.setBold(bold);
            f.setFontHeightInPoints((short) size);
            f.setFontName("맑은 고딕"); // ← 추가
            s.setFont(f);
            s.setVerticalAlignment(VerticalAlignment.CENTER);
            s.setBorderTop(BorderStyle.THIN);
            s.setBorderBottom(BorderStyle.THIN);
            s.setBorderLeft(BorderStyle.THIN);
            s.setBorderRight(BorderStyle.THIN);
            return s;
        }

        // ← setFontName("맑은 고딕") 추가
        private static CellStyle plain(XSSFWorkbook wb, boolean bold, int size) {
            XSSFCellStyle s = wb.createCellStyle();
            XSSFFont f = wb.createFont();
            f.setBold(bold);
            f.setFontHeightInPoints((short) size);
            f.setFontName("맑은 고딕"); // ← 추가
            s.setFont(f);
            s.setVerticalAlignment(VerticalAlignment.CENTER);
            return s;
        }

        private static void align(CellStyle s, HorizontalAlignment ha) {
            ((XSSFCellStyle) s).setAlignment(ha);
        }

        private static void tint(CellStyle s, byte r, byte g, byte b) {
            ((XSSFCellStyle) s).setFillForegroundColor(
                    new XSSFColor(new byte[]{r, g, b}, null));
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

    private void setCTTextCenter(XSSFTextBox box) {
        CTTextBody txBody = box.getCTShape().getTxBody();
        if (txBody == null) return;
        for (CTTextParagraph para : txBody.getPList()) {
            CTTextParagraphProperties pPr = para.isSetPPr()
                    ? para.getPPr()
                    : para.addNewPPr();
            pPr.setAlgn(STTextAlignType.CTR);
        }
    }
}