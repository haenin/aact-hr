package com.aact.schedule.service;

import com.aact.schedule.ScheduleRecordService;
import com.aact.schedule.dto.ScheduleRecordDto;
import com.aact.schedule.repository.ScheduleRecordRepository;
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
import java.util.Map;

import static com.aact.schedule.fixture.ScheduleTestFixture.*;
import static org.assertj.core.api.Assertions.assertThat;

/**
 * <pre>
 * Class Name: ScheduleRecordServiceTest
 * Description: 스케줄 엑셀 생성 및 서비스 통합 테스트
 *
 * History
 * 2026/05/14 (혜원) 엑셀 생성 및 파싱 테스트 성공
 * 2026/05/15 (혜원) 테스트 픽스처 분리 (ScheduleTestFixture), 데이터셋 3종 추가
 * </pre>
 */
@SpringBootTest
@Transactional
@DisplayName("스케줄 서비스 통합 테스트")
class ScheduleExcelServiceTest {

    @Autowired private ScheduleRecordService scheduleRecordService;
    @Autowired private ScheduleRecordRepository scheduleRecordRepository;

    @BeforeEach
    void setUp() {
        scheduleRecordRepository.deleteAll();
    }

    // ────────────────────────────────────────────────────────────────
    // Dataset A — 정상 케이스 (계약 & 교육품질팀 · 10인)
    // ────────────────────────────────────────────────────────────────

    @Test
    @DisplayName("[A] DB에 저장된 스케줄 데이터로 엑셀 파일 생성")
    void A_generateExcel_success() throws Exception {
        scheduleRecordRepository.saveAll(datasetA());

        byte[] excel = scheduleRecordService.generateExcel(requestFor(DEPT_A, FLIGHT_CODE_A, APPROVERS_A));

        assertThat(excel).isNotNull();
        assertThat(excel.length).isGreaterThan(0);

        Files.createDirectories(Paths.get("build"));
        Files.write(Paths.get("build/test-schedule-A.xlsx"), excel);
        System.out.println("엑셀 저장 완료: build/test-schedule-A.xlsx");
    }

    @Test
    @DisplayName("[A] 생성된 엑셀이 정상적으로 파싱되고 직원 이름이 존재한다")
    void A_generateExcelAndParse() throws Exception {
        scheduleRecordRepository.saveAll(datasetA());

        byte[] excel = scheduleRecordService.generateExcel(requestFor(DEPT_A, FLIGHT_CODE_A, APPROVERS_A));

        try (Workbook wb = new XSSFWorkbook(new ByteArrayInputStream(excel))) {
            assertThat(wb.getNumberOfSheets()).isGreaterThan(0);

            Sheet sheet = wb.getSheetAt(0);
            System.out.println("시트명: " + sheet.getSheetName());
            printSheet(sheet);

            assertThat(findName(sheet, "추순호")).isTrue();
        }
    }

    @Test
    @DisplayName("[A] 연월 + 부서로 스케줄 조회 시 10명이 순번 순으로 반환된다")
    void A_getByYearMonthAndDepartment() {
        scheduleRecordRepository.saveAll(datasetA());

        List<ScheduleRecordDto.Response> result =
                scheduleRecordService.getByYearMonthAndDepartment(YEAR_MONTH, DEPT_A);

        assertThat(result).hasSize(10);
        assertThat(result.get(0).getEmployeeName()).isEqualTo("추순호");
        assertThat(result.get(0).getSeq()).isEqualTo(1);

        System.out.println("=== [A] 조회 결과 ===");
        result.forEach(r -> System.out.println(r.getSeq() + ". " + r.getEmployeeName() + " | days: " + r.getDays()));
    }

    // ────────────────────────────────────────────────────────────────
    // Dataset B — 교대근무 케이스 (화물운영팀 · 6인)
    // ────────────────────────────────────────────────────────────────

    @Test
    @DisplayName("[B] 교대근무 패턴 데이터로 엑셀 파일 생성")
    void B_generateExcel_success() throws Exception {
        scheduleRecordRepository.saveAll(datasetB());

        byte[] excel = scheduleRecordService.generateExcel(requestFor(DEPT_B, FLIGHT_CODE_B, APPROVERS_B));

        assertThat(excel).isNotNull();
        assertThat(excel.length).isGreaterThan(0);

        Files.write(Paths.get("build/test-schedule-B.xlsx"), excel);
        System.out.println("엑셀 저장 완료: build/test-schedule-B.xlsx");
    }

    @Test
    @DisplayName("[B] 연차·반차·대체휴무 코드가 엑셀에 정상 출력된다")
    void B_generateExcel_containsSpecialCodes() throws Exception {
        scheduleRecordRepository.saveAll(datasetB());

        byte[] excel = scheduleRecordService.generateExcel(requestFor(DEPT_B, FLIGHT_CODE_B, APPROVERS_B));

        try (Workbook wb = new XSSFWorkbook(new ByteArrayInputStream(excel))) {
            Sheet sheet = wb.getSheetAt(0);
            boolean foundYeonja = false, foundBanja = false, foundDaeche = false;

            for (Row row : sheet) {
                for (Cell cell : row) {
                    if (cell.getCellType() != CellType.STRING) continue;
                    String v = cell.getStringCellValue();
                    if (v.equals("연")) foundYeonja = true;
                    if (v.equals("반")) foundBanja  = true;
                    if (v.equals("대")) foundDaeche = true;
                }
            }
            assertThat(foundYeonja).isTrue();
            assertThat(foundBanja).isTrue();
            assertThat(foundDaeche).isTrue();
        }
    }

    @Test
    @DisplayName("[B] 연월 + 부서로 스케줄 조회 시 6명이 순번 순으로 반환된다")
    void B_getByYearMonthAndDepartment() {
        scheduleRecordRepository.saveAll(datasetB());

        List<ScheduleRecordDto.Response> result =
                scheduleRecordService.getByYearMonthAndDepartment(YEAR_MONTH, DEPT_B);

        assertThat(result).hasSize(6);
        assertThat(result.get(0).getEmployeeName()).isEqualTo("강민준");

        System.out.println("=== [B] 조회 결과 ===");
        result.forEach(r -> System.out.println(r.getSeq() + ". " + r.getEmployeeName()));
    }

    // ────────────────────────────────────────────────────────────────
    // Dataset C — 스트레스 케이스 (수출입지원팀 · 15인)
    // ────────────────────────────────────────────────────────────────

    @Test
    @DisplayName("[C] 15인 데이터로 엑셀 파일 생성")
    void C_generateExcel_success() throws Exception {
        scheduleRecordRepository.saveAll(datasetC());

        byte[] excel = scheduleRecordService.generateExcel(requestFor(DEPT_C, FLIGHT_CODE_C, APPROVERS_C));

        assertThat(excel).isNotNull();
        assertThat(excel.length).isGreaterThan(0);

        Files.write(Paths.get("build/test-schedule-C.xlsx"), excel);
        System.out.println("엑셀 저장 완료: build/test-schedule-C.xlsx");
    }

    @Test
    @DisplayName("[C] X/9 복합 코드가 엑셀 셀에 정상 출력된다")
    void C_generateExcel_containsComplexCode() throws Exception {
        scheduleRecordRepository.saveAll(datasetC());

        byte[] excel = scheduleRecordService.generateExcel(requestFor(DEPT_C, FLIGHT_CODE_C, APPROVERS_C));

        try (Workbook wb = new XSSFWorkbook(new ByteArrayInputStream(excel))) {
            Sheet sheet = wb.getSheetAt(0);
            boolean foundXSlash9 = false;

            for (Row row : sheet) {
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.STRING && "X/9".equals(cell.getStringCellValue())) {
                        foundXSlash9 = true;
                        break;
                    }
                }
                if (foundXSlash9) break;
            }
            assertThat(foundXSlash9).isTrue();
        }
    }

    @Test
    @DisplayName("[C] 15인 데이터 — 시트 마지막 행 인덱스가 충분히 크다")
    void C_generateExcel_hasEnoughRows() throws Exception {
        scheduleRecordRepository.saveAll(datasetC());

        byte[] excel = scheduleRecordService.generateExcel(requestFor(DEPT_C, FLIGHT_CODE_C, APPROVERS_C));

        try (Workbook wb = new XSSFWorkbook(new ByteArrayInputStream(excel))) {
            // rowIdx: 제목(0~1) + 빈행(2) + 헤더(3~4) + 데이터(5~19) = 최소 19
            assertThat(wb.getSheetAt(0).getLastRowNum()).isGreaterThanOrEqualTo(19);
            System.out.println("[C] 마지막 행 인덱스: " + wb.getSheetAt(0).getLastRowNum());
        }
    }

    // ────────────────────────────────────────────────────────────────
    // 공통 — CRUD
    // ────────────────────────────────────────────────────────────────

    @Test
    @DisplayName("스케줄 단건 등록 후 id로 조회하면 저장된 days 값이 일치한다")
    void createAndGet() {
        ScheduleRecordDto.Response saved = scheduleRecordService.create(
                ScheduleRecordDto.Request.builder()
                        .applyYearMonth(YEAR_MONTH).department(DEPT_A).flightCode(FLIGHT_CODE_A)
                        .employeeName("테스트직원").seq(11)
                        .days(Map.ofEntries(
                                Map.entry(1,"X"), Map.entry(2,"9"),
                                Map.entry(3,"연"), Map.entry(4,"반"), Map.entry(5,"X/9")))
                        .offDays(10).usedOff(2.0).usedAnnual(1.0).remainAnnual(9.0)
                        .build());

        ScheduleRecordDto.Response found = scheduleRecordService.getById(saved.getId());

        assertThat(found.getEmployeeName()).isEqualTo("테스트직원");
        assertThat(found.getDays().get(3)).isEqualTo("연");
        assertThat(found.getDays().get(5)).isEqualTo("X/9");
        System.out.println("=== 저장된 days === " + found.getDays());
    }

    @Test
    @DisplayName("스케줄 수정 시 변경된 날짜값이 정상적으로 반영된다")
    void update() {
        ScheduleRecordDto.Response saved = scheduleRecordService.create(
                ScheduleRecordDto.Request.builder()
                        .applyYearMonth(YEAR_MONTH).department(DEPT_A).flightCode(FLIGHT_CODE_A)
                        .employeeName("수정테스트").seq(12).days(Map.of(1,"9"))
                        .build());

        ScheduleRecordDto.Response updated = scheduleRecordService.update(
                saved.getId(),
                ScheduleRecordDto.Request.builder()
                        .applyYearMonth(YEAR_MONTH).department(DEPT_A).flightCode(FLIGHT_CODE_A)
                        .employeeName("수정테스트").seq(12).days(Map.of(1,"X", 2,"연"))
                        .build());

        assertThat(updated.getDays().get(1)).isEqualTo("X");
        assertThat(updated.getDays().get(2)).isEqualTo("연");
        System.out.println("=== 수정 후 days === " + updated.getDays());
    }

    @Test
    @DisplayName("스케줄 단건 삭제 후 해당 id로 조회하면 존재하지 않는다")
    void delete() {
        ScheduleRecordDto.Response saved = scheduleRecordService.create(
                ScheduleRecordDto.Request.builder()
                        .applyYearMonth(YEAR_MONTH).department(DEPT_A).flightCode(FLIGHT_CODE_A)
                        .employeeName("삭제테스트").seq(13)
                        .build());

        scheduleRecordService.delete(saved.getId());

        assertThat(scheduleRecordRepository.findById(saved.getId())).isEmpty();
        System.out.println("=== 삭제 완료 ===");
    }

    // ────────────────────────────────────────────────────────────────
    // 헬퍼
    // ────────────────────────────────────────────────────────────────

    private boolean findName(Sheet sheet, String name) {
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().contains(name)) {
                    System.out.println(name + " 데이터 발견!");
                    return true;
                }
            }
        }
        return false;
    }

    private void printSheet(Sheet sheet) {
        System.out.println("===== 엑셀 내용 =====");
        for (Row row : sheet) {
            StringBuilder sb = new StringBuilder();
            for (Cell cell : row) {
                sb.append(switch (cell.getCellType()) {
                    case STRING  -> cell.getStringCellValue();
                    case NUMERIC -> String.valueOf((int) cell.getNumericCellValue());
                    case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
                    default -> "";
                }).append("\t");
            }
            System.out.println(sb);
        }
    }
}