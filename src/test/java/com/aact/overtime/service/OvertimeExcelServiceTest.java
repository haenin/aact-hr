package com.aact.overtime.service;

import com.aact.overtime.dto.OvertimeRecordDto;
import com.aact.overtime.dto.DepartmentSummaryDto;
import com.aact.overtime.entity.OvertimeRecord;
import com.aact.overtime.entity.DepartmentSummary;
import com.aact.overtime.repository.OvertimeRecordRepository;
import com.aact.overtime.repository.DepartmentSummaryRepository;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.transaction.annotation.Transactional;

import java.io.ByteArrayInputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.LocalTime;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * <pre>
 * Class Name: OvertimeExcelServiceTest
 * Description: 엑셀 생성 및 파싱 테스트
 *
 * History
 * 2026/05/14 (혜원) 엑셀 생성 및 파싱 테스트 성공
 * </pre>
 */
@SpringBootTest
@Transactional
@DisplayName("시간외근무 엑셀 통합 테스트")
class OvertimeExcelServiceTest {

    @Autowired private OvertimeRecordService overtimeRecordService;
    @Autowired private DepartmentSummaryService departmentSummaryService;
    @Autowired private ExcelSheetService excelSheetService;
    @Autowired private OvertimeRecordRepository overtimeRecordRepository;
    @Autowired private DepartmentSummaryRepository departmentSummaryRepository;

    private static final String YEAR_MONTH = "2026-05";
    private static final String DEPARTMENT = "항공기운송지원팀";

    @BeforeEach
    void setUp() {
        // ── DepartmentSummary (표지 집계표) 가상 데이터
        departmentSummaryRepository.saveAll(List.of(
            makeSummary("조업부",       20.0, 5.0, 3.0,  22.0, 6.0, 2.0,  "업무증가"),
            makeSummary("항공기운송지원팀", 15.0, 3.0, 2.0,  18.0, 4.0, 3.0,  "항공편 증가"),
            makeSummary("관리부",       10.0, 2.0, 1.0,  12.0, 2.0, 1.0,  null)
        ));

        // ── OvertimeRecord (세부내역) 가상 데이터
        overtimeRecordRepository.saveAll(List.of(
            makeRecord("김보라",  LocalDate.of(2026, 5, 1),  "09:00", "18:00", "09:00", "20:30", 2.5, 0.0, 0.0, 0.0),
            makeRecord("김보라",  LocalDate.of(2026, 5, 8),  "09:00", "18:00", "09:00", "21:00", 3.0, 0.0, 0.0, 0.0),
            makeRecord("김보라",  LocalDate.of(2026, 5, 15), "09:00", "18:00", "09:00", "22:00", 4.0, 1.0, 0.0, 0.0),
            makeRecord("모은서",  LocalDate.of(2026, 5, 2),  "09:00", "18:00", "09:00", "20:00", 2.0, 0.0, 0.0, 0.0),
            makeRecord("모은서",  LocalDate.of(2026, 5, 9),  "09:00", "18:00", "09:00", "23:00", 5.0, 0.0, 0.0, 0.0),
            makeRecord("남다정",  LocalDate.of(2026, 5, 3),  "13:00", "22:00", "13:00", "22:00", 0.0, 4.0, 0.0, 0.0),
            makeRecord("남다정",  LocalDate.of(2026, 5, 10), "13:00", "22:00", "13:00", "23:30", 1.5, 4.0, 0.0, 0.0)
        ));
    }

    // ────────────────────────────────────────────────────────────
    // 주 테스트
    // ────────────────────────────────────────────────────────────

    @Test
    @DisplayName("DB에 저장된 데이터로 시간외근무수당 엑셀 파일 생성")
    void generateExcel() throws Exception {
        byte[] excel = excelSheetService.generateExcel(
                YEAR_MONTH,
                List.of("담당", "과장", "상무", "부사장"));

        assertThat(excel).isNotNull();
        assertThat(excel.length).isGreaterThan(0);

        Path path = Paths.get("build/test-overtime.xlsx");
        Files.write(path, excel);
        System.out.println("엑셀 저장 완료: " + path.toAbsolutePath());
    }

    @Test
    @DisplayName("생성된 엑셀에 표지 시트와 세부내역 시트가 모두 존재한다")
    void generateExcelAndVerifySheets() throws Exception {
        byte[] excel = excelSheetService.generateExcel(
                YEAR_MONTH,
                List.of("담당", "과장", "상무", "부사장"));

        try (Workbook workbook = new XSSFWorkbook(new ByteArrayInputStream(excel))) {
            int sheetCount = workbook.getNumberOfSheets();
            assertThat(sheetCount).isGreaterThan(0);

            System.out.println("===== 시트 목록 =====");
            for (int i = 0; i < sheetCount; i++) {
                System.out.println("시트" + (i + 1) + ": " + workbook.getSheetAt(i).getSheetName());
            }

            // 표지 시트 확인
            boolean hasCover = false;
            boolean hasDetail = false;
            for (int i = 0; i < sheetCount; i++) {
                String name = workbook.getSheetAt(i).getSheetName();
                if (name.contains("표지")) hasCover = true;
                else hasDetail = true;
            }
            assertThat(hasCover).isTrue();
            assertThat(hasDetail).isTrue();
        }
    }

    @Test
    @DisplayName("생성된 엑셀 세부내역 시트에 직원 이름이 존재한다")
    void generateExcelAndParse() throws Exception {
        byte[] excel = excelSheetService.generateExcel(
                YEAR_MONTH,
                List.of("담당", "과장", "상무", "부사장"));

        Path path = Paths.get("build/test-overtime.xlsx");
        Files.write(path, excel);
        System.out.println("엑셀 저장 완료: " + path.toAbsolutePath());

        try (Workbook workbook = new XSSFWorkbook(new ByteArrayInputStream(excel))) {

            // 세부내역 시트 찾기
            Sheet detailSheet = null;
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                if (!workbook.getSheetAt(i).getSheetName().contains("표지")) {
                    detailSheet = workbook.getSheetAt(i);
                    break;
                }
            }
            assertThat(detailSheet).isNotNull();

            System.out.println("시트명: " + detailSheet.getSheetName());
            System.out.println("===== 엑셀 내용 =====");

            boolean foundKimBora = false;
            for (Row row : detailSheet) {
                StringBuilder sb = new StringBuilder();
                for (Cell cell : row) {
                    String value = switch (cell.getCellType()) {
                        case STRING  -> cell.getStringCellValue();
                        case NUMERIC -> String.valueOf((int) cell.getNumericCellValue());
                        case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
                        default -> "";
                    };
                    sb.append(value).append("\t");
                    if (cell.getCellType() == CellType.STRING &&
                            cell.getStringCellValue().contains("김보라")) {
                        foundKimBora = true;
                        System.out.println("김보라 데이터 발견!");
                    }
                }
                System.out.println(sb);
            }
            assertThat(foundKimBora).isTrue();
        }
    }

    @Test
    @DisplayName("직원별 소계가 정상적으로 계산된다")
    void getSubtotal() {
        List<OvertimeRecordDto.Subtotal> subtotals =
                overtimeRecordService.getSubtotalByDepartment(YEAR_MONTH, DEPARTMENT);

        assertThat(subtotals).isNotEmpty();
        System.out.println("===== 직원별 소계 =====");
        subtotals.forEach(s -> System.out.println(
                s.getEmployeeName() + " | 연장: " + s.getExtensionHours()
                + " | 야간: " + s.getNightHours()
                + " | 휴일: " + s.getHolidayHours()));

        // 김보라 소계 검증 (2.5 + 3.0 + 4.0 = 9.5)
        subtotals.stream()
                .filter(s -> "김보라".equals(s.getEmployeeName()))
                .findFirst()
                .ifPresent(s -> assertThat(s.getExtensionHours()).isEqualTo(9.5));
    }

    @Test
    @DisplayName("부서 전체 합계가 정상적으로 계산된다")
    void getTotal() {
        OvertimeRecordDto.Total total =
                overtimeRecordService.getTotalByDepartment(YEAR_MONTH, DEPARTMENT);

        assertThat(total).isNotNull();
        System.out.println("===== 부서 합계 =====");
        System.out.println(total.getDepartment()
                + " | 연장: " + total.getExtensionHours()
                + " | 야간: " + total.getNightHours());
    }

    @Test
    @DisplayName("부서별 집계표 조회가 정상적으로 동작한다")
    void getDepartmentSummary() {
        List<DepartmentSummaryDto.Response> summaries =
                departmentSummaryService.getByYearMonth(YEAR_MONTH);

        assertThat(summaries).hasSize(3);
        System.out.println("===== 부서별 집계표 =====");
        summaries.forEach(s -> System.out.println(
                s.getDepartment()
                + " | 전월연장: " + s.getPrevExtensionHours()
                + " | 당월연장: " + s.getCurrExtensionHours()
                + " | 증감: "    + s.getDiffExtensionHours()
                + " | 사유: "    + s.getChangeReason()));
    }

    // ────────────────────────────────────────────────────────────
    // 헬퍼
    // ────────────────────────────────────────────────────────────

    private DepartmentSummary makeSummary(String dept,
                                          double prevExt, double prevNight, double prevHol,
                                          double currExt, double currNight, double currHol,
                                          String reason) {
        return DepartmentSummary.builder()
                .applyYearMonth(YEAR_MONTH)
                .department(dept)
                .prevExtensionHours(prevExt)
                .prevNightHours(prevNight)
                .prevHolidayHours(prevHol)
                .currExtensionHours(currExt)
                .currNightHours(currNight)
                .currHolidayHours(currHol)
                .diffExtensionHours(currExt - prevExt)
                .diffNightHours(currNight - prevNight)
                .diffHolidayHours(currHol - prevHol)
                .changeReason(reason)
                .build();
    }

    private OvertimeRecord makeRecord(String name, LocalDate date,
                                       String schStart, String schEnd,
                                       String actStart, String actEnd,
                                       double ext, double night, double hol, double holExt) {
        return OvertimeRecord.builder()
                .applyYearMonth(YEAR_MONTH)
                .department(DEPARTMENT)
                .employeeName(name)
                .workDate(date)
                .scheduledStart(LocalTime.parse(schStart))
                .scheduledEnd(LocalTime.parse(schEnd))
                .actualStart(LocalTime.parse(actStart))
                .actualEnd(LocalTime.parse(actEnd))
                .extensionHours(ext)
                .nightHours(night)
                .holidayHours(hol)
                .holidayExtensionHours(holExt)
                .build();
    }
}
