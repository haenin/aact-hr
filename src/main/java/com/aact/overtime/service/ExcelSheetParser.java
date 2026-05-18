package com.aact.overtime.service;

import com.aact.overtime.dto.ExcelSheetDto;
import com.aact.overtime.dto.ExcelSheetType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Component;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

@Component
public class ExcelSheetParser {

    @SuppressWarnings("unchecked")
    public List<ExcelSheetDto.SheetWrapper<?>> parse(InputStream is) throws Exception {
        List<ExcelSheetDto.SheetWrapper<?>> result = new ArrayList<>();

        try (Workbook wb = new XSSFWorkbook(is)) {
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                Sheet sheet = wb.getSheetAt(i);
                String sheetName = sheet.getSheetName();

                boolean isCover = sheetName.contains("표지");
                ExcelSheetType type = ExcelSheetType.from(sheetName);

                if (isCover) {
                    List<ExcelSheetDto.CoverRow> data = parseCoverSheet(sheet);
                    List<String> approvers = parseApprovers(sheet);
                    result.add(ExcelSheetDto.SheetWrapper.<ExcelSheetDto.CoverRow>builder()
                            .sheetName(sheetName)
                            .isCover(true)
                            .type(type)
                            .data(data)
                            .approvers(approvers)
                            .build());
                } else {
                    List<ExcelSheetDto.DetailRow> data = parseDetailSheet(sheet);
                    List<String> approvers = parseApprovers(sheet);
                    result.add(ExcelSheetDto.SheetWrapper.<ExcelSheetDto.DetailRow>builder()
                            .sheetName(sheetName)
                            .isCover(false)
                            .type(type)
                            .data(data)
                            .approvers(approvers)
                            .build());
                }
            }
        }
        return result;
    }

    // ────────────────────────────────────────────────────────────
    // 결재란 파싱 (Drawing 텍스트박스)
    // ────────────────────────────────────────────────────────────
    private List<String> parseApprovers(Sheet sheet) {
        List<String> approvers = new ArrayList<>();
        if (!(sheet instanceof XSSFSheet xssfSheet)) return approvers;

        XSSFDrawing drawing = xssfSheet.getDrawingPatriarch();
        if (drawing == null) return approvers;

        for (XSSFShape shape : drawing.getShapes()) {
            // XSSFTextBox → 저장 후 재파싱하면 XSSFSimpleShape으로 역직렬화됨
            if (shape instanceof XSSFSimpleShape simpleShape) {
                try {
                    String text = simpleShape.getText().trim();
                    if (!text.isBlank()) {
                        approvers.add(text);
                    }
                } catch (Exception ignored) {}
            }
        }
        return approvers;
    }
    // ────────────────────────────────────────────────────────────
    // 표지 시트 파싱
    // ────────────────────────────────────────────────────────────
    private List<ExcelSheetDto.CoverRow> parseCoverSheet(Sheet sheet) {
        List<ExcelSheetDto.CoverRow> rows = new ArrayList<>();
        boolean dataStarted = false;

        for (Row row : sheet) {
            String firstCell = strVal(row.getCell(0));

            if ("부서명".equals(firstCell)) {
                dataStarted = true;
                continue;
            }
            if (!dataStarted) continue;
            if (firstCell == null || firstCell.isBlank()) continue;
            if (firstCell.contains("상기와") || firstCell.contains("에이에이씨티")) break;
            if ("연장".equals(firstCell) || "야간".equals(firstCell) || "휴일".equals(firstCell)) continue;

            rows.add(ExcelSheetDto.CoverRow.builder()
                    .departmentName(firstCell)
                    .prevExtension(numVal(row.getCell(1)))
                    .prevNight(numVal(row.getCell(2)))
                    .prevHoliday(numVal(row.getCell(3)))
                    .currExtension(numVal(row.getCell(4)))
                    .currNight(numVal(row.getCell(5)))
                    .currHoliday(numVal(row.getCell(6)))
                    .diffExtension(diff(numVal(row.getCell(4)), numVal(row.getCell(1))))
                    .diffNight(diff(numVal(row.getCell(5)), numVal(row.getCell(2))))
                    .diffHoliday(diff(numVal(row.getCell(6)), numVal(row.getCell(3))))
                    .changeReason(strVal(row.getCell(33)))
                    .build());
        }
        return rows;
    }

    // ────────────────────────────────────────────────────────────
    // 세부내역 시트 파싱
    // ────────────────────────────────────────────────────────────
    private List<ExcelSheetDto.DetailRow> parseDetailSheet(Sheet sheet) {
        List<ExcelSheetDto.DetailRow> rows = new ArrayList<>();
        boolean dataStarted = false;

        String currentDept = null;
        String currentName = null;

        for (Row row : sheet) {
            String col0 = strVal(row.getCell(0));
            String col1 = strVal(row.getCell(1));
            String col2 = strVal(row.getCell(2));

            if ("부서".equals(col0)) {
                dataStarted = true;
                continue;
            }
            if (!dataStarted) continue;
            if (col1 != null && col1.contains("합계")) break;

            if ("소계".equals(col1)) {
                rows.add(ExcelSheetDto.DetailRow.builder()
                        .department(currentDept)
                        .employeeName(currentName)
                        .isSubtotal(true)
                        .extensionHours(numVal(row.getCell(7)))
                        .nightHours(numVal(row.getCell(8)))
                        .holidayHours(numVal(row.getCell(9)))
                        .holidayExtensionHours(numVal(row.getCell(10)))
                        .build());
                continue;
            }

            if (col0 != null && !col0.isBlank()) currentDept = col0;
            if (col1 != null && !col1.isBlank()) currentName = col1;
            if (col2 == null || col2.isBlank()) continue;

            rows.add(ExcelSheetDto.DetailRow.builder()
                    .department(currentDept)
                    .employeeName(currentName)
                    .workDate(col2)
                    .scheduledStart(strVal(row.getCell(3)))
                    .scheduledEnd(strVal(row.getCell(4)))
                    .actualStart(strVal(row.getCell(5)))
                    .actualEnd(strVal(row.getCell(6)))
                    .extensionHours(numVal(row.getCell(7)))
                    .nightHours(numVal(row.getCell(8)))
                    .holidayHours(numVal(row.getCell(9)))
                    .holidayExtensionHours(numVal(row.getCell(10)))
                    .isSubtotal(false)
                    .build());
        }
        return rows;
    }

    private String strVal(Cell cell) {
        if (cell == null) return null;
        return switch (cell.getCellType()) {
            case STRING  -> cell.getStringCellValue().trim();
            case NUMERIC -> String.valueOf((long) cell.getNumericCellValue());
            case FORMULA -> {
                try { yield cell.getStringCellValue().trim(); }
                catch (Exception e) { yield String.valueOf(cell.getNumericCellValue()); }
            }
            default -> null;
        };
    }

    private Double numVal(Cell cell) {
        if (cell == null) return 0.0;
        return switch (cell.getCellType()) {
            case NUMERIC -> cell.getNumericCellValue();
            case FORMULA -> {
                try { yield cell.getNumericCellValue(); }
                catch (Exception e) { yield 0.0; }
            }
            case STRING  -> {
                try { yield Double.parseDouble(cell.getStringCellValue()); }
                catch (Exception e) { yield 0.0; }
            }
            default -> 0.0;
        };
    }

    private Double diff(Double curr, Double prev) {
        return (curr != null ? curr : 0.0) - (prev != null ? prev : 0.0);
    }
}