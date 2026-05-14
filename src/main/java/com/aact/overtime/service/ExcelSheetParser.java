package com.aact.overtime.service;

import com.aact.overtime.dto.SheetDto;
import com.aact.overtime.dto.SheetType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * 시간외근무수당 엑셀 파일을 SheetDto 배열로 파싱
 *
 * 반환 구조:
 * List<SheetDto.SheetWrapper<?>>
 *  ├── isCover=true  → data: List<CoverRow>
 *  └── isCover=false → data: List<DetailRow>
 */
@Component
public class ExcelSheetParser {

    /**
     * 엑셀 파일 전체를 시트 배열로 변환
     */
    @SuppressWarnings("unchecked")
    public List<SheetDto.SheetWrapper<?>> parse(InputStream is) throws Exception {
        List<SheetDto.SheetWrapper<?>> result = new ArrayList<>();

        try (Workbook wb = new XSSFWorkbook(is)) {
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                Sheet sheet = wb.getSheetAt(i);
                String sheetName = sheet.getSheetName();

                boolean isCover = sheetName.contains("표지");
                SheetType type  = SheetType.from(sheetName);

                if (isCover) {
                    List<SheetDto.CoverRow> data = parseCoverSheet(sheet);
                    result.add(SheetDto.SheetWrapper.<SheetDto.CoverRow>builder()
                            .sheetName(sheetName)
                            .isCover(true)
                            .type(type)
                            .data(data)
                            .build());
                } else {
                    List<SheetDto.DetailRow> data = parseDetailSheet(sheet);
                    result.add(SheetDto.SheetWrapper.<SheetDto.DetailRow>builder()
                            .sheetName(sheetName)
                            .isCover(false)
                            .type(type)
                            .data(data)
                            .build());
                }
            }
        }
        return result;
    }

    // ────────────────────────────────────────────────────────────
    // 표지 시트 파싱
    // 부서명 / 전월(연장·야간·휴일) / 당월(연장·야간·휴일) / 증감 / 사유
    // 실제 데이터 행: "총 합계" 위의 부서명 행들
    // ────────────────────────────────────────────────────────────
    private List<SheetDto.CoverRow> parseCoverSheet(Sheet sheet) {
        List<SheetDto.CoverRow> rows = new ArrayList<>();
        boolean dataStarted = false;

        for (Row row : sheet) {
            String firstCell = strVal(row.getCell(0));

            // "부서명" 헤더 다음 행부터 데이터
            if ("부서명".equals(firstCell)) {
                dataStarted = true;
                continue;
            }
            if (!dataStarted) continue;

            // 빈 행 또는 서명 구간 스킵
            if (firstCell == null || firstCell.isBlank()) continue;
            if (firstCell.contains("상기와") || firstCell.contains("에이에이씨티")) break;

            // 헤더 보조행 ("연장","야간","휴일" 반복) 스킵
            if ("연장".equals(firstCell) || "야간".equals(firstCell) || "휴일".equals(firstCell)) continue;

            rows.add(SheetDto.CoverRow.builder()
                    .departmentName(firstCell)
                    .prevExtension(numVal(row.getCell(1)))
                    .prevNight(numVal(row.getCell(2)))
                    .prevHoliday(numVal(row.getCell(3)))
                    .currExtension(numVal(row.getCell(4)))
                    .currNight(numVal(row.getCell(5)))
                    .currHoliday(numVal(row.getCell(6)))
                    // 증감은 수식 결과 or 직접 계산
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
    // 부서 / 성명 / 일자 / 예정출근·퇴근 / 실출근·퇴근 / 연장·야간·휴일·휴일연장
    // ────────────────────────────────────────────────────────────
    private List<SheetDto.DetailRow> parseDetailSheet(Sheet sheet) {
        List<SheetDto.DetailRow> rows = new ArrayList<>();
        boolean dataStarted = false;

        String currentDept = null;
        String currentName = null;

        for (Row row : sheet) {
            String col0 = strVal(row.getCell(0)); // 부서
            String col1 = strVal(row.getCell(1)); // 성명
            String col2 = strVal(row.getCell(2)); // 일자

            // "부서" 헤더 행 다음부터 데이터
            if ("부서".equals(col0)) {
                dataStarted = true;
                continue;
            }
            if (!dataStarted) continue;

            // 합계 행 스킵 (시트 끝)
            if (col1 != null && col1.contains("합계")) break;

            // 소계 행
            if ("소계".equals(col1)) {
                rows.add(SheetDto.DetailRow.builder()
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

            // 부서/성명 carry-forward (병합셀 대응)
            if (col0 != null && !col0.isBlank()) currentDept = col0;
            if (col1 != null && !col1.isBlank()) currentName = col1;

            // 일자가 없으면 빈 행
            if (col2 == null || col2.isBlank()) continue;

            rows.add(SheetDto.DetailRow.builder()
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

    // ─── 셀 값 추출 유틸 ───────────────────────────────────────

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
