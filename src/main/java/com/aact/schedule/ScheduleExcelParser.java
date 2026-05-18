package com.aact.schedule;

import com.aact.schedule.dto.ScheduleRecordDto;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Component;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * 스케줄 엑셀 파싱
 *
 * 구조:
 *  row 0~1  : 제목 + 결재란 (셀 방식)
 *  row 2    : 빈 행
 *  row 3    : 헤더 1행 (순번 | 일자 | 1~31일 | 휴무일수 | 사용휴무 | 사용연차 | 잔여연차)
 *  row 4    : 헤더 2행 (성명 | 요일)
 *  row 5~   : 데이터 행
 */
@Component
public class ScheduleExcelParser {

    public List<ScheduleRecordDto.ParsedRow> parse(InputStream is,
                                                   String applyYearMonth,
                                                   String department) throws Exception {
        List<ScheduleRecordDto.ParsedRow> result = new ArrayList<>();

        try (Workbook wb = new XSSFWorkbook(is)) {
            Sheet sheet = wb.getSheetAt(0);

            // 헤더 2행(row 3)에서 마지막 날짜 열 인덱스 파악
            Row headerRow = sheet.getRow(3);
            int lastDayCol = 1; // 최소 1(성명)
            if (headerRow != null) {
                for (int c = 2; c <= 32; c++) {
                    Cell cell = headerRow.getCell(c);
                    if (cell == null) break;
                    String val = strVal(cell);
                    if (val == null || val.isBlank()) break;
                    try {
                        Integer.parseInt(val);
                        lastDayCol = c;
                    } catch (NumberFormatException e) {
                        break;
                    }
                }
            }
            int lastDay = lastDayCol - 1; // 실제 마지막 날짜 (1~31)
            int tailStart = lastDayCol + 1; // 휴무일수 시작 열

            // row 5부터 데이터
            for (int r = 5; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                String seqStr = strVal(row.getCell(0));
                String name   = strVal(row.getCell(1));

                if (seqStr == null || seqStr.isBlank()) continue;
                int seq;
                try {
                    seq = Integer.parseInt(seqStr.trim());
                } catch (NumberFormatException e) {
                    continue;
                }

                List<String> days = new ArrayList<>();
                for (int d = 1; d <= lastDay; d++) {
                    days.add(nvl(strVal(row.getCell(1 + d))));
                }

                result.add(ScheduleRecordDto.ParsedRow.builder()
                        .applyYearMonth(applyYearMonth)
                        .department(department)
                        .seq(seq)
                        .employeeName(nvl(name))
                        .days(days)
                        .offDays(strVal(row.getCell(tailStart)))
                        .usedOff(strVal(row.getCell(tailStart + 1)))
                        .usedAnnual(strVal(row.getCell(tailStart + 2)))
                        .remainAnnual(strVal(row.getCell(tailStart + 3)))
                        .build());
            }
        }
        return result;
    }

    public List<String> parseApprovers(InputStream is) throws Exception {
        List<String> approvers = new ArrayList<>();
        try (Workbook wb = new XSSFWorkbook(is)) {
            Sheet sheet = wb.getSheetAt(0);
            if (!(sheet instanceof XSSFSheet xssfSheet)) return approvers;

            // 결재란은 셀 방식 (row 0~1, approvalStartCol=32~)
            Row r0 = sheet.getRow(0);
            if (r0 == null) return approvers;

            for (int c = 33; c <= 36; c++) {
                Cell cell = r0.getCell(c);
                if (cell == null) continue;
                String val = strVal(cell);
                if (val != null && !val.isBlank()) {
                    approvers.add(val);
                }
            }
        }
        return approvers;
    }

    private String strVal(Cell cell) {
        if (cell == null) return null;
        return switch (cell.getCellType()) {
            case STRING  -> cell.getStringCellValue().trim();
            case NUMERIC -> {
                double v = cell.getNumericCellValue();
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

    private String nvl(String s) {
        return s != null ? s : "";
    }
}