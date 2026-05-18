package com.aact.overtime.service;

import com.aact.overtime.dto.ExcelSheetDto;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.transaction.annotation.Transactional;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

@SpringBootTest
@Transactional
@DisplayName("시간외근무 엑셀 파싱 테스트")
class OvertimeExcelSheetParserTest {

    @Autowired private ExcelSheetParser excelSheetParser;

    @Test
    @DisplayName("실제 엑셀 파일을 파싱하면 표지/세부내역 데이터가 정상 출력된다")
    void parse_realExcel_printAll() throws Exception {
        try (InputStream is = getClass().getResourceAsStream("/fixture/test-overtime-A.xlsx")) {

            assertThat(is).as("fixture/test-overtime-A.xlsx 파일이 test/resources에 없습니다").isNotNull();
            byte[] bytes = is.readAllBytes();

            // when
            List<ExcelSheetDto.SheetWrapper<?>> sheets =
                    excelSheetParser.parse(new ByteArrayInputStream(bytes));

            // then
            assertThat(sheets).isNotEmpty();

            for (ExcelSheetDto.SheetWrapper<?> sheet : sheets) {

                if (sheet.isCover()) {
                    // ── 표지
                    System.out.println("\n");
                    System.out.println("╔══════════════════════════════════════════════════════════════╗");
                    System.out.println("║                        [표지] 부서별 집계표                        ║");
                    System.out.println("╠══════════════════════════════════════════════════════════════╣");
                    System.out.printf( "║ 시트명: %-54s║%n", sheet.getSheetName());
                    System.out.println("╠══════════════════════════════════════════════════════════════╣");

                    List<String> approvers = sheet.getApprovers();
                    System.out.println("║ [결재란]                                                       ║");
                    if (approvers == null || approvers.isEmpty()) {
                        System.out.println("║   (없음)                                                      ║");
                    } else {
                        approvers.forEach(a ->
                                System.out.printf("║   - %-57s║%n", a.replace("\n", " / "))
                        );
                    }

                    System.out.println("╠══════════════╦═══════════════════╦═══════════════════╦════════╣");
                    System.out.println("║    부서명     ║  전월(연장/야간/휴일)  ║  당월(연장/야간/휴일)  ║  사유  ║");
                    System.out.println("╠══════════════╬═══════════════════╬═══════════════════╬════════╣");

                    List<ExcelSheetDto.CoverRow> rows = (List<ExcelSheetDto.CoverRow>) sheet.getData();
                    if (rows.isEmpty()) {
                        System.out.println("║                        (파싱된 데이터 없음)                       ║");
                    } else {
                        for (ExcelSheetDto.CoverRow row : rows) {
                            System.out.printf("║ %-12s ║ %5.1f / %5.1f / %5.1f ║ %5.1f / %5.1f / %5.1f ║ %-6s ║%n",
                                    row.getDepartmentName(),
                                    row.getPrevExtension(), row.getPrevNight(), row.getPrevHoliday(),
                                    row.getCurrExtension(), row.getCurrNight(), row.getCurrHoliday(),
                                    row.getChangeReason() != null ? row.getChangeReason() : "-");
                        }
                    }
                    System.out.println("╚══════════════╩═══════════════════╩═══════════════════╩════════╝");
                    System.out.println("  → 집계 " + rows.size() + "행 / 결재란 " + (approvers != null ? approvers.size() : 0) + "개 파싱");

                } else {
                    // ── 세부내역
                    System.out.println("\n");
                    System.out.println("╔══════════════════════════════════════════════════════════════════════╗");
                    System.out.printf( "║                [세부내역] %-45s║%n", sheet.getSheetName());
                    System.out.println("╠══════════════════════════════════════════════════════════════════════╣");

                    List<String> approvers = sheet.getApprovers();
                    System.out.println("║ [결재란]                                                              ║");
                    if (approvers == null || approvers.isEmpty()) {
                        System.out.println("║   (없음)                                                             ║");
                    } else {
                        approvers.forEach(a ->
                                System.out.printf("║   - %-67s║%n", a.replace("\n", " / "))
                        );
                    }

                    System.out.println("╠══════════╦══════╦══════════╦══════════════╦══════════════╦══════╦══════╦══════╣");
                    System.out.println("║   부서   ║ 성명 ║   일자   ║  예정(출~퇴) ║  실제(출~퇴) ║ 연장 ║ 야간 ║ 휴일 ║");
                    System.out.println("╠══════════╬══════╬══════════╬══════════════╬══════════════╬══════╬══════╬══════╣");

                    List<ExcelSheetDto.DetailRow> rows = (List<ExcelSheetDto.DetailRow>) sheet.getData();
                    if (rows.isEmpty()) {
                        System.out.println("║                            (파싱된 데이터 없음)                              ║");
                    } else {
                        for (ExcelSheetDto.DetailRow row : rows) {
                            if (row.isSubtotal()) {
                                System.out.println("╠══════════╩══════╩══════════╩══════════════╩══════════════╩══════╩══════╩══════╣");
                                System.out.printf("║ ★ 소계  %-6s %47s연장 %4.1f / 야간 %4.1f / 휴일 %4.1f   ║%n",
                                        nvl(row.getEmployeeName()), "",
                                        row.getExtensionHours(), row.getNightHours(), row.getHolidayHours());
                                System.out.println("╠══════════╦══════╦══════════╦══════════════╦══════════════╦══════╦══════╦══════╣");
                            } else {
                                System.out.printf("║ %-8s ║ %-4s ║ %-8s ║ %5s ~ %5s ║ %5s ~ %5s ║ %4.1f ║ %4.1f ║ %4.1f ║%n",
                                        nvl(row.getDepartment()),
                                        nvl(row.getEmployeeName()),
                                        nvl(row.getWorkDate()),
                                        nvl(row.getScheduledStart()), nvl(row.getScheduledEnd()),
                                        nvl(row.getActualStart()),    nvl(row.getActualEnd()),
                                        row.getExtensionHours(), row.getNightHours(), row.getHolidayHours());
                            }
                        }
                    }
                    System.out.println("╚══════════╩══════╩══════════╩══════════════╩══════════════╩══════╩══════╩══════╝");
                    System.out.println("  → 총 " + rows.size() + "행 파싱 (소계 포함) / 결재란 " + (approvers != null ? approvers.size() : 0) + "개 파싱");
                }
            }
        }
    }

    private String nvl(String s) {
        return s != null ? s : "-";
    }
}