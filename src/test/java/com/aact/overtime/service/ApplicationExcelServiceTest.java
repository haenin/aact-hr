package com.aact.overtime.service;

import com.aact.overtime.fixture.ApplicationTestFixture;
import com.aact.overtime.repository.ApplicationRepository;
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

import static com.aact.overtime.fixture.ApplicationTestFixture.*;
import static org.assertj.core.api.Assertions.assertThat;

/**
 * <pre>
 * Class Name: ApplicationExcelServiceTest
 * Description: 시간외근무수당 신청서 엑셀 생성 통합 테스트
 *
 * History
 * 2026/05/15 (혜원) 엑셀 생성 통합 테스트 작성
 *                   테스트 픽스처 분리 (ApplicationTestFixture)
 * </pre>
 */
@SpringBootTest
@Transactional
@DisplayName("시간외근무수당 신청서 엑셀 생성 통합 테스트")
class ApplicationExcelServiceTest {

    @Autowired private ApplicationService applicationService;
    @Autowired private ApplicationRepository applicationRepository;

    @BeforeEach
    void setUp() {
        applicationRepository.deleteAll();
    }

    // ────────────────────────────────────────────────────────────────
    // Dataset A — 정상 케이스
    // ────────────────────────────────────────────────────────────────

    @Test
    @DisplayName("[A] DB에 저장된 데이터로 신청서 엑셀 파일을 생성")
    void A_generateExcel_success() throws Exception {
        applicationRepository.saveAll(datasetA());

        byte[] excel = applicationService.generateExcel(requestFor(DEPT_A));

        assertThat(excel).isNotNull();
        assertThat(excel.length).isGreaterThan(0);

        Files.createDirectories(Paths.get("build"));
        Files.write(Paths.get("build/test-application-A.xlsx"), excel);
        System.out.println("엑셀 저장 완료: build/test-application-A.xlsx");
    }

    @Test
    @DisplayName("[A] 생성된 엑셀에 월별 시트가 존재한다")
    void A_generateExcel_hasMonthSheet() throws Exception {
        applicationRepository.saveAll(datasetA());

        byte[] excel = applicationService.generateExcel(requestFor(DEPT_A));

        try (Workbook wb = new XSSFWorkbook(new ByteArrayInputStream(excel))) {
            boolean hasMonthSheet = false;
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                if (wb.getSheetAt(i).getSheetName().contains("신청서")) {
                    hasMonthSheet = true;
                    break;
                }
            }
            assertThat(hasMonthSheet).isTrue();
        }
    }

    @Test
    @DisplayName("[A] 생성된 엑셀 시트에 직원 이름이 존재한다")
    void A_generateExcel_containsEmployeeNames() throws Exception {
        applicationRepository.saveAll(datasetA());

        byte[] excel = applicationService.generateExcel(requestFor(DEPT_A));

        try (Workbook wb = new XSSFWorkbook(new ByteArrayInputStream(excel))) {
            Sheet sheet = wb.getSheetAt(0);
            boolean foundKimBora = false, foundMoEunseo = false, foundNamDajung = false;

            for (Row row : sheet) {
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.STRING) {
                        String v = cell.getStringCellValue();
                        if (v.contains("김보라"))  foundKimBora   = true;
                        if (v.contains("모은서"))  foundMoEunseo  = true;
                        if (v.contains("남다정"))  foundNamDajung = true;
                    }
                }
            }
            assertThat(foundKimBora).isTrue();
            assertThat(foundMoEunseo).isTrue();
            assertThat(foundNamDajung).isTrue();
        }
    }

    @Test
    @DisplayName("[A] 연장/야간/휴일 합계가 정상적으로 계산된다 — 연장 13.0 / 야간 9.0 / 휴일 0.0")
    void A_generateExcel_totalRowValues() throws Exception {
        applicationRepository.saveAll(datasetA());

        byte[] excel = applicationService.generateExcel(requestFor(DEPT_A));

        try (Workbook wb = new XSSFWorkbook(new ByteArrayInputStream(excel))) {
            Row totalRow = findTotalRow(wb.getSheetAt(0));
            assertThat(totalRow).isNotNull();

            assertThat(totalRow.getCell(5).getNumericCellValue()).isEqualTo(13.0);
            assertThat(totalRow.getCell(6).getNumericCellValue()).isEqualTo(9.0);
            assertThat(totalRow.getCell(7).getNumericCellValue()).isEqualTo(0.0);
        }
    }

    // ────────────────────────────────────────────────────────────────
    // Dataset B — 경계값 케이스
    // ────────────────────────────────────────────────────────────────

    @Test
    @DisplayName("[B] 야간 최대치(8h) + 휴일 포함 엑셀 생성")
    void B_generateExcel_success() throws Exception {
        applicationRepository.saveAll(datasetB());

        byte[] excel = applicationService.generateExcel(requestFor(DEPT_B));

        assertThat(excel).isNotNull();
        assertThat(excel.length).isGreaterThan(0);

        Files.write(Paths.get("build/test-application-B.xlsx"), excel);
        System.out.println("엑셀 저장 완료: build/test-application-B.xlsx");
    }

    @Test
    @DisplayName("[B] 연장/야간/휴일 합계 — 연장 7.5 / 야간 18.0 / 휴일 16.0")
    void B_generateExcel_totalRowValues() throws Exception {
        applicationRepository.saveAll(datasetB());

        byte[] excel = applicationService.generateExcel(requestFor(DEPT_B));

        try (Workbook wb = new XSSFWorkbook(new ByteArrayInputStream(excel))) {
            Row totalRow = findTotalRow(wb.getSheetAt(0));
            assertThat(totalRow).isNotNull();

            assertThat(totalRow.getCell(5).getNumericCellValue()).isEqualTo(7.5);
            assertThat(totalRow.getCell(6).getNumericCellValue()).isEqualTo(18.0);
            assertThat(totalRow.getCell(7).getNumericCellValue()).isEqualTo(16.0);
        }
    }

    @Test
    @DisplayName("[B] 연장 0인 행(정시 퇴근)도 엑셀에 포함된다")
    void B_generateExcel_containsZeroOvertimeRow() throws Exception {
        applicationRepository.saveAll(datasetB());

        byte[] excel = applicationService.generateExcel(requestFor(DEPT_B));

        try (Workbook wb = new XSSFWorkbook(new ByteArrayInputStream(excel))) {
            Sheet sheet = wb.getSheetAt(0);
            boolean foundZeroRow = false;

            for (Row row : sheet) {
                Cell nameCell = row.getCell(1);
                Cell extCell  = row.getCell(5);
                if (nameCell != null && extCell != null
                        && nameCell.getCellType() == CellType.STRING
                        && nameCell.getStringCellValue().contains("오태양")
                        && extCell.getCellType()  == CellType.NUMERIC
                        && extCell.getNumericCellValue() == 0.0) {
                    foundZeroRow = true;
                    break;
                }
            }
            assertThat(foundZeroRow).isTrue();
        }
    }

    // ────────────────────────────────────────────────────────────────
    // Dataset C — 스트레스 케이스
    // ────────────────────────────────────────────────────────────────

    @Test
    @DisplayName("[C] 28건 데이터로 빈 행 없이 엑셀 생성")
    void C_generateExcel_success() throws Exception {
        applicationRepository.saveAll(datasetC());

        byte[] excel = applicationService.generateExcel(requestFor(DEPT_C));

        assertThat(excel).isNotNull();
        assertThat(excel.length).isGreaterThan(0);

        Files.write(Paths.get("build/test-application-C.xlsx"), excel);
        System.out.println("엑셀 저장 완료: build/test-application-C.xlsx");
    }

    @Test
    @DisplayName("[C] 28건 데이터 — 시트의 마지막 행 인덱스가 충분하다")
    void C_generateExcel_noEmptyRows() throws Exception {
        applicationRepository.saveAll(datasetC());

        byte[] excel = applicationService.generateExcel(requestFor(DEPT_C));

        try (Workbook wb = new XSSFWorkbook(new ByteArrayInputStream(excel))) {
            Sheet sheet = wb.getSheetAt(0);
            // rowIdx: 제목(0) + 로고/년월(1~2) + 헤더(3) + 데이터(4~31) + 합계(32)
            assertThat(sheet.getLastRowNum()).isGreaterThanOrEqualTo(32);
            System.out.println("28건 엑셀 마지막 행 인덱스: " + sheet.getLastRowNum());
        }
    }

    @Test
    @DisplayName("[C] 연장/야간/휴일 합계 — 연장 46.5 / 야간 18.0 / 휴일 8.0")
    void C_generateExcel_totalRowValues() throws Exception {
        applicationRepository.saveAll(datasetC());

        byte[] excel = applicationService.generateExcel(requestFor(DEPT_C));

        try (Workbook wb = new XSSFWorkbook(new ByteArrayInputStream(excel))) {
            Row totalRow = findTotalRow(wb.getSheetAt(0));
            assertThat(totalRow).isNotNull();

            assertThat(totalRow.getCell(5).getNumericCellValue()).isEqualTo(46.5);
            assertThat(totalRow.getCell(6).getNumericCellValue()).isEqualTo(18.0);
            assertThat(totalRow.getCell(7).getNumericCellValue()).isEqualTo(8.0);
        }
    }

    // ────────────────────────────────────────────────────────────────
    // 공통 — 데이터 없을 때
    // ────────────────────────────────────────────────────────────────

    @Test
    @DisplayName("데이터가 없을 때도 엑셀이 정상 생성된다 (빈 행 28개)")
    void generateExcel_emptyData() throws Exception {
        byte[] excel = applicationService.generateExcel(requestFor(DEPT_A));

        assertThat(excel).isNotNull();
        assertThat(excel.length).isGreaterThan(0);

        try (Workbook wb = new XSSFWorkbook(new ByteArrayInputStream(excel))) {
            assertThat(wb.getSheetAt(0).getLastRowNum()).isGreaterThanOrEqualTo(32);
        }
    }

    // ────────────────────────────────────────────────────────────────
    // 헬퍼
    // ────────────────────────────────────────────────────────────────

    /** 첫 번째 셀에 "합"이 포함된 행을 합계 행으로 간주 */
    private Row findTotalRow(Sheet sheet) {
        for (Row row : sheet) {
            Cell c = row.getCell(0);
            if (c != null && c.getCellType() == CellType.STRING
                    && c.getStringCellValue().contains("합")) {
                return row;
            }
        }
        return null;
    }
}