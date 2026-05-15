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

            sheet.setDisplayGridlines(false); // 시트의 기본 눈금선을 보이지 않게 설정
            wb.write(out);
            return out.toByteArray();
        }
    }

    // ── [row 0] 제목
    private int createTitleRow(XSSFSheet sheet, StyleSet styles,
                               String dept, int rowIdx) {
        Row r = sheet.createRow(rowIdx);
        r.setHeightInPoints(80);

        // [핵심 수정] 0~8번 셀을 생성하되, 테두리가 없는 스타일(noBorder)을 입힙니다.
        // styles.title을 입히면 bordered 설정 때문에 선이 생깁니다.
        for (int i = 0; i <= 8; i++) {
            Cell cell = r.createCell(i);
            cell.setCellStyle(styles.noBorder);
        }

        // 셀 병합
        sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx, 0, 8));

        // 0번 셀에만 제목 전용 스타일(폰트 크기+정렬)을 입힙니다.
        // 단, styles.title에 테두리 설정이 있다면 아래처럼 새로 정의한 스타일을 쓰는게 좋습니다.
        Cell c = r.getCell(0);
        c.setCellValue("시간 외 근무수당 신청서 (" + dept + ")");

        // 테두리 없는 제목 전용 스타일 적용 (아래 StyleSet 수정 참고)
        c.setCellStyle(styles.titleNoBorder);

        return ++rowIdx;
    }

    /**
     * [row 1~2]
     *  col 0~7 : 로고(row1 col0~6 병합), 나머지 빈칸
     *  col 8   : 결재란 Drawing (근무사유 열 위에만 배치)
     *
     *  결재란 구조 (col 8 내부를 dx 오프셋으로 3분할):
     *   ┌──────┬──────┬──────┐
     *   │  결재 │  담당 │  과장 │  ← row1 (헤더)
     *   │      ├──────┼──────┤
     *   │      │      │      │  ← row2 (서명칸)
     *   └──────┴──────┴──────┘
     *
     *  col 8 너비 = 10000 units
     *  EMU 환산: 10000 * 914400 / 1024 ≈ 8,929,687 → 편의상 2,600,000 EMU 사용
     *  결재 레이블 : 1/4 = 650,000 EMU
     *  담당/과장   : 각 3/8 = 975,000 EMU
     */
    private int createLogoAndApprovalBlock(XSSFSheet sheet, StyleSet styles,
                                           List<String> approvers, int rowIdx,
                                           XSSFWorkbook wb) {
        Row row1 = sheet.createRow(rowIdx);
        row1.setHeightInPoints(40);
        sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx, 0, 6));
        for (int col = 0; col <= 6; col++) {
            row1.createCell(col).setCellStyle(styles.noBorder);
        }

        // ← drawing 딱 한 번만 생성
        XSSFDrawing drawing = sheet.createDrawingPatriarch();

        // 이미지 삽입
        try {
            var is = getClass().getResourceAsStream("/img.png");
            if (is != null) {
                byte[] bytes = is.readAllBytes();
                int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
                is.close();

                // ← 위에서 만든 drawing 재사용
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

        // 결재란도 같은 drawing 재사용 (새로 createDrawingPatriarch() 호출 X)
        List<String> headers = (approvers != null && !approvers.isEmpty())
                ? approvers : List.of("담당", "대리");

        // --- EMU 계산 로직 (동적 변경) ---
        int totalEmu  = 3_000_000; // 전체 너비를 조금 더 넓게 조정 (선택 사항)
        int offsetEmu = 0;   // 왼쪽 여백
        int labelEmu  = 400_000;   // "결재" 박스 너비

        int count = headers.size();
        int eachEmu = (totalEmu - labelEmu) / count; // 결재자 수만큼 균등 배분

        // 1. "결재" 레이블 생성 (rowIdx ~ rowIdx + 2 높이)
        XSSFClientAnchor labelAnchor = drawing.createAnchor(
                offsetEmu, 0,
                offsetEmu + labelEmu, 0,
                8, rowIdx,
                8, rowIdx + 2
        );
        XSSFTextBox labelBox = drawing.createTextbox(labelAnchor);
        applyBoxStyle(labelBox, wb, "결\n재", true, true); // 공통 스타일 적용

        // 2. 결재자별 헤더 및 서명칸 생성
        for (int i = 0; i < count; i++) {
            // x축 위치 계산: 시작점 + 결재레이블 + (순번 * 개별너비)
            int currentStartX = offsetEmu + labelEmu + (i * eachEmu);
            int currentEndX   = currentStartX + eachEmu;

            // 헤더 (상단: 담당/대리/과장)
            XSSFClientAnchor hAnchor = drawing.createAnchor(
                    currentStartX, 0, currentEndX, 0,
                    8, rowIdx,
                    8, rowIdx + 1
            );
            XSSFTextBox hBox = drawing.createTextbox(hAnchor);
            applyBoxStyle(hBox, wb, headers.get(i), true, true);

            // 서명칸 (하단: 빈칸)
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
     * 텍스트박스 스타일 중복 코드를 줄이기 위한 헬퍼 메서드
     */
    private void applyBoxStyle(XSSFTextBox box, XSSFWorkbook wb, String text, boolean isBold, boolean hasFill) {
        box.setLineStyleColor(0, 0, 0);
        box.setLineWidth(0.5); // [수정] 0.75에서 0.5로 더 얇게 조정

        if (hasFill) {
            box.setFillColor(217, 217, 217);
        } else {
            box.setNoFill(true);
        }

        XSSFFont font = wb.createFont();
        font.setBold(isBold);
        font.setFontHeightInPoints((short) 9);

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

    // ── 테이블 헤더 (줄바꿈 없이 한 줄, 근무사유 중앙정렬)
    private int createTableHeader(XSSFSheet sheet, StyleSet styles, int rowIdx) {
        Row r = sheet.createRow(rowIdx);
        r.setHeightInPoints(24);
        setCell(r, 0, "NO",          styles.header);
        setCell(r, 1, "일 자",        styles.header);
        setCell(r, 2, "성 명",        styles.header);
        setCell(r, 3, "예정근무시간",   styles.header);
        setCell(r, 4, "실근무시간",    styles.header);
        setCell(r, 5, "연장근무시간",  styles.header);  // 줄바꿈 제거
        setCell(r, 6, "야간근무시간",  styles.header);  // 줄바꿈 제거
        setCell(r, 7, "휴일근무시간",  styles.header);  // 줄바꿈 제거
        setCell(r, 8, "근무사유",  styles.header);  // 중앙정렬 (header와 동일)
        return ++rowIdx;
    }

    // ── 데이터 행 (근무사유 중앙정렬)
    private void createDataRow(XSSFSheet sheet, StyleSet styles,
                               OvertimePayApplication rec, int rowIdx) {
        Row r = sheet.createRow(rowIdx);
        // 데이터 행의 높이를 설정
        r.setHeightInPoints(20);
        // 1. 날짜 가공 로직
        String originalDate = rec.getWorkDate(); // "2026-05-02"
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

        // 2. 셀에 데이터 채우기
        setCell(r, 0, String.valueOf(rec.getNo()), styles.dataCenter);
        setCell(r, 1, formattedDate, styles.dataCenter); // 가공된 "05월 02일"

        setCell   (r, 2, rec.getEmployeeName()  != null ? rec.getEmployeeName()  : "",  styles.dataCenter);
        setCell   (r, 3, rec.getScheduledTime() != null ? rec.getScheduledTime() : "",  styles.dataCenter);
        setCell   (r, 4, rec.getActualTime()    != null ? rec.getActualTime()    : "",  styles.dataCenter);
        setNumCell(r, 5, rec.getExtensionHours(),  styles.dataCenter);
        setNumCell(r, 6, rec.getNightHours(),      styles.dataCenter);
        setNumCell(r, 7, rec.getHolidayHours(),    styles.dataCenter);
        setCell   (r, 8, rec.getWorkReason()    != null ? rec.getWorkReason()    : "",  styles.dataCenter);
    }

    // ── 합계 행 수정
    private void createTotalRow(XSSFSheet sheet, StyleSet styles,
                                double ext, double night, double holiday, int rowIdx) {
        Row r = sheet.createRow(rowIdx);
        r.setHeightInPoints(22);

        for (int i = 0; i <= 8; i++) {
            Cell c = r.createCell(i);
            // 기본적으로 테두리가 있는 스타일 적용
            if (i <= 4) {
                c.setCellStyle(styles.totalLabel); // 합계 글자 쪽 배경색 있는 스타일
            } else {
                c.setCellStyle(styles.header);     // 숫자 및 빈칸 쪽 테두리 스타일
            }
        }

        // 2. 0~4번 컬럼 병합
        sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx, 0, 4));

        // 3. 값 다시 세팅 (이미 위에서 스타일을 입혔으므로 값만 넣으면 됨)
        r.getCell(0).setCellValue("합    계");
        r.getCell(5).setCellValue(ext);
        r.getCell(6).setCellValue(night);
        r.getCell(7).setCellValue(holiday);
        // r.getCell(8)은 이미 위 루프에서 styles.header가 적용되어 테두리가 생깁니다.
    }

    /**
     * 서명란 — col 8 아래, 합계 행 다음 행부터
     *  예시:  신 청 자 :  홍길동   (인)
     *         팀    장 :  김팀장   (인)
     */
    private void createSignerRows(XSSFSheet sheet, StyleSet styles,
                                  List<ApplicationDto.Signer> signers, int rowIdx) {
        if (signers == null || signers.isEmpty()) return;

        int startRow = rowIdx + 1;
        Workbook wb = sheet.getWorkbook();

        // 1. 이름용 볼드 폰트 생성
        XSSFFont boldFont = (XSSFFont) wb.createFont();
        boldFont.setBold(true);
        boldFont.setFontHeightInPoints((short) 11);

        // 2. (인) 및 기본 텍스트용 일반 폰트 생성
        XSSFFont plainFont = (XSSFFont) wb.createFont();
        plainFont.setBold(false);
        plainFont.setFontHeightInPoints((short) 11);

        for (int i = 0; i < signers.size(); i++) {
            ApplicationDto.Signer signer = signers.get(i);
            Row r = sheet.createRow(startRow + i);
            r.setHeightInPoints(36);

            // 왼쪽 빈 셀들 스타일 적용
            for (int col = 0; col <= 6; col++) {
                r.createCell(col).setCellStyle(styles.noBorder);
            }

            // --- 역할 라벨 처리 (예: "신청자" -> "신 청 자") ---
            Cell roleCell = r.createCell(7);
            String rawRole = signer.getRole() != null ? signer.getRole() : "";

            // 글자들 사이에 공백 1칸씩 삽입 (글자가 2자면 "A B", 3자면 "A B C")
            String displayRole = String.join(" ", rawRole.split(""));

            // 만약 2글자일 때 3글자 너비처럼 더 벌리고 싶다면 아래 로직 사용
            if (rawRole.length() == 2) {
                displayRole = rawRole.charAt(0) + "    " + rawRole.charAt(1); // 공백 4칸
            }

            roleCell.setCellValue(displayRole + " :");
            roleCell.setCellStyle(styles.signerRole);

            // --- 이름 및 (인) 처리 (이름만 볼드) ---
            Cell nameCell = r.createCell(8);
            XSSFRichTextString richText = new XSSFRichTextString();

            // 이름 부분 (Bold)
            richText.append("  " + (signer.getName() != null ? signer.getName() : ""), boldFont);

            // 간격 및 (인) (Plain)
            // %-60s 같은 방식보다 공백 문자를 직접 넉넉히 넣는게 엑셀에서 더 안정적입니다.
            richText.append("                                                                 (인)", plainFont);

            nameCell.setCellValue(richText);
            nameCell.setCellStyle(styles.signerLine);
        }
    }

    // ────────────────────────────────────────────────
    // StyleSet
    // ────────────────────────────────────────────────
    private static class StyleSet {
        final CellStyle title,titleNoBorder, approvalLabel, approvalHeader, approvalSign;
        final CellStyle logo, yearMonth, header, totalLabel;
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

            // 헤더: 줄바꿈 OFF, 중앙정렬
            header = bordered(wb, true, 9);
            align(header, HorizontalAlignment.CENTER);
            header.setWrapText(false);
            tint(header, (byte)217, (byte)217, (byte)217);

            totalLabel = bordered(wb, true, 10);
            align(totalLabel, HorizontalAlignment.CENTER);
            tint(totalLabel, (byte)230, (byte)230, (byte)230);

            dataCenter = bordered(wb, false, 8);
            align(dataCenter, HorizontalAlignment.CENTER);
            // 날짜 컬럼 전용 스타일 추가
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

        private CellStyle bordered(XSSFWorkbook wb, boolean bold, int size) {
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

        private CellStyle plain(XSSFWorkbook wb, boolean bold, int size) {
            XSSFCellStyle s = wb.createCellStyle();
            XSSFFont f = wb.createFont();
            f.setBold(bold); f.setFontHeightInPoints((short) size);
            s.setFont(f);
            s.setVerticalAlignment(VerticalAlignment.CENTER);
            return s;
        }

        private void align(CellStyle s, HorizontalAlignment ha) {
            ((XSSFCellStyle) s).setAlignment(ha);
        }

        private void tint(CellStyle s, byte r, byte g, byte b) {
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

    /**
     * XSSFTextBox 내부 텍스트를 가로 중앙정렬.
     * XSSFSimpleShape.setHorizontalAlignment() 은 POI 버전에 따라 없을 수 있으므로
     * CTTextBody를 직접 조작한다.
     */
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