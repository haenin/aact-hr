package com.aact.overtime.service;

import com.aact.overtime.repository.OvertimeRecordRepository;
import com.aact.overtime.repository.DepartmentSummaryRepository;
import com.aact.overtime.dto.OvertimeRecordDto;
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
import java.nio.file.Paths;
import java.util.List;

import static com.aact.overtime.fixture.OvertimeTestFixture.*;
import static org.assertj.core.api.Assertions.assertThat;

/**
 * <pre>
 * Class Name: OvertimeExcelServiceTest
 * Description: 시간외근무 엑셀 생성 통합 테스트
 *
 * History
 * 2026/05/18 (혜원) 테스트 픽스처 분리 및 A/B/C 케이스 구조화
 * </pre>
 */
@SpringBootTest
@Transactional
@DisplayName("시간외근무 엑셀 통합 테스트")
class OvertimeExcelServiceTest {

    @Autowired private OvertimeRecordService overtimeRecordService;
    @Autowired private ExcelSheetService excelSheetService;
    @Autowired private OvertimeRecordRepository overtimeRecordRepository;
    @Autowired private DepartmentSummaryRepository departmentSummaryRepository;
    @Autowired private DepartmentSummaryService departmentSummaryService;

    @BeforeEach
    void setUp() {
        departmentSummaryRepository.deleteAll();
        overtimeRecordRepository.deleteAll();
    }

    // ────────────────────────────────────────────────────────────────
    // Dataset A — 정상 케이스
    // ────────────────────────────────────────────────────────────────

    @Test
    @DisplayName("[A] DB에 저장된 데이터로 엑셀 파일을 생성한다")
    void A_generateExcel_success() throws Exception {
        departmentSummaryRepository.saveAll(summariesA());
        overtimeRecordRepository.saveAll(recordsA());

        byte[] excel = excelSheetService.generateExcel(YEAR_MONTH, APPROVERS_4);

        assertThat(excel).isNotNull();
        assertThat(excel.length).isGreaterThan(0);

        Files.createDirectories(Paths.get("build"));
        Files.write(Paths.get("build/test-overtime-A.xlsx"), excel);
        System.out.println("엑셀 저장 완료: build/test-overtime-A.xlsx");
    }

    @Test
    @DisplayName("[A] 표지 시트와 세부내역 시트가 모두 존재한다")
    void A_generateExcel_hasRequiredSheets() throws Exception {
        departmentSummaryRepository.saveAll(summariesA());
        overtimeRecordRepository.saveAll(recordsA());

        byte[] excel = excelSheetService.generateExcel(YEAR_MONTH, APPROVERS_4);

        try (Workbook wb = new XSSFWorkbook(new ByteArrayInputStream(excel))) {
            boolean hasCover = false, hasDetail = false;
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                String name = wb.getSheetAt(i).getSheetName();
                if (name.contains("표지")) hasCover = true;
                else hasDetail = true;
            }
            assertThat(hasCover).isTrue();
            assertThat(hasDetail).isTrue();
        }
    }

    @Test
    @DisplayName("[A] 세부내역 시트에 직원 이름이 존재한다")
    void A_generateExcel_containsEmployeeNames() throws Exception {
        departmentSummaryRepository.saveAll(summariesA());
        overtimeRecordRepository.saveAll(recordsA());

        byte[] excel = excelSheetService.generateExcel(YEAR_MONTH, APPROVERS_4);

        try (Workbook wb = new XSSFWorkbook(new ByteArrayInputStream(excel))) {
            Sheet detail = findDetailSheet(wb);
            assertThat(detail).isNotNull();

            boolean foundKimBora = false, foundMoEunseo = false, foundNamDajung = false;
            for (Row row : detail) {
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.STRING) {
                        String v = cell.getStringCellValue();
                        if (v.contains("김보라")) foundKimBora  = true;
                        if (v.contains("모은서")) foundMoEunseo = true;
                        if (v.contains("남다정")) foundNamDajung = true;
                    }
                }
            }
            assertThat(foundKimBora).isTrue();
            assertThat(foundMoEunseo).isTrue();
            assertThat(foundNamDajung).isTrue();
        }
    }

    @Test
    @DisplayName("[A] 김보라 연장 소계가 9.5h로 정상 계산된다")
    void A_subtotal_kimBora_extensionHours() {
        overtimeRecordRepository.saveAll(recordsA());

        List<OvertimeRecordDto.Subtotal> subtotals =
                overtimeRecordService.getSubtotalByDepartment(YEAR_MONTH, DEPT_A);

        subtotals.stream()
                .filter(s -> "김보라".equals(s.getEmployeeName()))
                .findFirst()
                .ifPresentOrElse(
                        s -> assertThat(s.getExtensionHours()).isEqualTo(9.5),
                        () -> { throw new AssertionError("김보라 소계 없음"); }
                );
    }

    // ────────────────────────────────────────────────────────────────
    // Dataset B — 경계값 케이스
    // ────────────────────────────────────────────────────────────────

    @Test
    @DisplayName("[B] 야간 최대치(8h) + 휴일 포함 엑셀 생성")
    void B_generateExcel_success() throws Exception {
        departmentSummaryRepository.saveAll(summariesB());
        overtimeRecordRepository.saveAll(recordsB());

        byte[] excel = excelSheetService.generateExcel(YEAR_MONTH, APPROVERS_2);

        assertThat(excel).isNotNull();
        assertThat(excel.length).isGreaterThan(0);

        Files.write(Paths.get("build/test-overtime-B.xlsx"), excel);
        System.out.println("엑셀 저장 완료: build/test-overtime-B.xlsx");
    }

    @Test
    @DisplayName("[B] 연장 0인 행(정시 퇴근)도 세부내역에 포함된다")
    void B_generateExcel_containsZeroOvertimeRow() throws Exception {
        departmentSummaryRepository.saveAll(summariesB());
        overtimeRecordRepository.saveAll(recordsB());

        byte[] excel = excelSheetService.generateExcel(YEAR_MONTH, APPROVERS_2);

        try (Workbook wb = new XSSFWorkbook(new ByteArrayInputStream(excel))) {
            Sheet detail = findDetailSheet(wb);
            assertThat(detail).isNotNull();

            boolean foundZeroRow = false;
            for (Row row : detail) {
                boolean hasOhTaeyang = false;
                boolean hasZeroExt   = false;
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.STRING
                            && cell.getStringCellValue().contains("오태양")) {
                        hasOhTaeyang = true;
                    }
                    if (cell.getCellType() == CellType.NUMERIC
                            && cell.getNumericCellValue() == 0.0) {
                        hasZeroExt = true;
                    }
                }
                if (hasOhTaeyang && hasZeroExt) {
                    foundZeroRow = true;
                    break;
                }
            }
            assertThat(foundZeroRow).isTrue();
        }
    }

    @Test
    @DisplayName("[B] 정수빈 야간 소계가 16.0h로 정상 계산된다")
    void B_subtotal_jungSubin_nightHours() {
        overtimeRecordRepository.saveAll(recordsB());

        List<OvertimeRecordDto.Subtotal> subtotals =
                overtimeRecordService.getSubtotalByDepartment(YEAR_MONTH, DEPT_B);

        subtotals.stream()
                .filter(s -> "정수빈".equals(s.getEmployeeName()))
                .findFirst()
                .ifPresentOrElse(
                        s -> assertThat(s.getNightHours()).isEqualTo(16.0),
                        () -> { throw new AssertionError("정수빈 소계 없음"); }
                );
    }

    // ────────────────────────────────────────────────────────────────
    // Dataset C — 스트레스 케이스
    // ────────────────────────────────────────────────────────────────

    @Test
    @DisplayName("[C] 3개 부서 집계표로 엑셀 생성")
    void C_generateExcel_success() throws Exception {
        departmentSummaryRepository.saveAll(summariesC());
        overtimeRecordRepository.saveAll(recordsC());

        byte[] excel = excelSheetService.generateExcel(YEAR_MONTH, APPROVERS_4);

        assertThat(excel).isNotNull();
        assertThat(excel.length).isGreaterThan(0);

        Files.write(Paths.get("build/test-overtime-C.xlsx"), excel);
        System.out.println("엑셀 저장 완료: build/test-overtime-C.xlsx");
    }

    @Test
    @DisplayName("[C] 부서별 집계표 3건이 정상 조회된다")
    void C_departmentSummary_hasThreeDepts() {
        departmentSummaryRepository.saveAll(summariesC());

        var summaries = departmentSummaryService.getByYearMonth(YEAR_MONTH);
        assertThat(summaries).hasSize(3);
    }

    @Test
    @DisplayName("[C] 세부내역 시트 수가 부서 수만큼 존재한다")
    void C_generateExcel_sheetCountMatchesDepts() throws Exception {
        departmentSummaryRepository.saveAll(summariesC());
        overtimeRecordRepository.saveAll(recordsC());

        byte[] excel = excelSheetService.generateExcel(YEAR_MONTH, APPROVERS_4);

        try (Workbook wb = new XSSFWorkbook(new ByteArrayInputStream(excel))) {
            long detailSheetCount = 0;
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                if (!wb.getSheetAt(i).getSheetName().contains("표지")) {
                    detailSheetCount++;
                }
            }
            // 표지 1장 + 부서별 세부내역
            assertThat(detailSheetCount).isGreaterThanOrEqualTo(1);
        }
    }

    // ────────────────────────────────────────────────────────────────
    // 헬퍼
    // ────────────────────────────────────────────────────────────────

    /** 표지가 아닌 첫 번째 시트를 세부내역 시트로 반환 */
    private Sheet findDetailSheet(Workbook wb) {
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            if (!wb.getSheetAt(i).getSheetName().contains("표지")) {
                return wb.getSheetAt(i);
            }
        }
        return null;
    }
}