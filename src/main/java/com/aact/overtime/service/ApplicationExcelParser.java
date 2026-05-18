package com.aact.overtime.service;

import com.aact.overtime.dto.ApplicationDto;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;
import org.apache.poi.xssf.usermodel.*;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * 시간외근무수당 신청서(관리부) 엑셀 파싱
 *
 * 구조:
 *  row 0     : 제목 (시간 외 근무수당 신청서(관리부)) + 결재란
 *  row 1~2   : AACT 로고 + 년도/월
 *  row 3     : 테이블 헤더 (NO | 일자 | 성명 | 예정근무시간 | 실근무시간 | 연장 | 야간 | 휴일 | 근무사유)
 *  row 4~    : 데이터 행 (최대 28행)
 *              - NO 있는 행 → 실제 데이터
 *              - NO 없고 ~ 있는 행 → 빈 행 스킵
 *              - 빈 행 → 스킵
 *  합계 행   : "합   계" 텍스트 → 스킵
 *  서명란    : 신청자/팀장/부서장 → 스킵
 */
@Component
public class ApplicationExcelParser {

    public List<ApplicationDto.Request> parse(InputStream is,
                                              String applyYearMonth,
                                              String department) throws Exception {
        List<ApplicationDto.Request> result = new ArrayList<>();

        try (Workbook wb = new XSSFWorkbook(is)) {
            Sheet sheet = wb.getSheetAt(0);

            boolean dataStarted = false;

            for (Row row : sheet) {
                String col0 = strVal(row.getCell(0)); // NO
                String col1 = strVal(row.getCell(1)); // 일자
                String col2 = strVal(row.getCell(2)); // 성명
                String col3 = strVal(row.getCell(3)); // 예정근무시간
                String col4 = strVal(row.getCell(4)); // 실근무시간
                String col5 = strVal(row.getCell(5)); // 연장
                String col6 = strVal(row.getCell(6)); // 야간
                String col7 = strVal(row.getCell(7)); // 휴일
                String col8 = strVal(row.getCell(8)); // 근무사유

                // "NO" 헤더 행 다음부터 데이터
                if ("NO".equals(col0)) {
                    dataStarted = true;
                    continue;
                }
                if (!dataStarted) continue;

                // 합계 행 → 종료
                if (col1 != null && col1.contains("합")) break;

                // NO가 숫자가 아닌 행 스킵 (빈 행, ~ 행, 서명란)
                if (col0 == null || col0.isBlank()) continue;
                int no;
                try {
                    no = Integer.parseInt(col0.trim());
                } catch (NumberFormatException e) {
                    continue;
                }

                // ~ 행 스킵 (빈 신청 행)
                if ("~".equals(col3) || "~".equals(col4)) continue;

                // 일자/성명 없는 행 스킵
                if (col1 == null || col1.isBlank()) continue;

                result.add(ApplicationDto.Request.builder()
                        .applyYearMonth(applyYearMonth)
                        .department(department)
                        .no(no)
                        .workDate(col1)
                        .employeeName(col2)
                        .scheduledTime(col3)
                        .actualTime(col4)
                        .extensionHours(numVal(row.getCell(5)))
                        .nightHours(numVal(row.getCell(6)))
                        .holidayHours(numVal(row.getCell(7)))
                        .workReason(col8)
                        .build());
            }
        }
        return result;
    }

    // ── 셀 값 추출
    private String strVal(Cell cell) {
        if (cell == null) return null;
        return switch (cell.getCellType()) {
            case STRING  -> cell.getStringCellValue().trim();
            case NUMERIC -> {
                double v = cell.getNumericCellValue();
                // 정수면 정수로 반환 ex) 1.0 → "1"
                yield v == Math.floor(v) ? String.valueOf((long) v) : String.valueOf(v);
            }
            case FORMULA -> {
                try { yield cell.getStringCellValue().trim(); }
                catch (Exception e) {
                    double v = cell.getNumericCellValue();
                    yield v == Math.floor(v) ? String.valueOf((long) v) : String.valueOf(v);
                }
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

    // parse() 메서드 반환타입을 감싸는 DTO가 없으면 일단 결재란만 출력용으로 추가
    public List<String> parseApprovers(InputStream is) throws Exception {
        List<String> approvers = new ArrayList<>();
        try (Workbook wb = new XSSFWorkbook(is)) {
            Sheet sheet = wb.getSheetAt(0);
            if (!(sheet instanceof XSSFSheet xssfSheet)) return approvers;

            XSSFDrawing drawing = xssfSheet.getDrawingPatriarch();
            if (drawing == null) return approvers;

            for (XSSFShape shape : drawing.getShapes()) {
                if (shape instanceof XSSFSimpleShape simpleShape) {
                    try {
                        String text = simpleShape.getText().trim();
                        if (!text.isBlank()) approvers.add(text);
                    } catch (Exception ignored) {}
                }
            }
        }
        return approvers;
    }
}
