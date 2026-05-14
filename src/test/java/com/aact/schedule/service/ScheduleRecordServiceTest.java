package com.aact.schedule.service;

import com.aact.schedule.dto.ScheduleRecordDto;
import com.aact.schedule.entity.ScheduleRecord;
import com.aact.schedule.repository.ScheduleRecordRepository;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.transaction.annotation.Transactional;

import java.util.List;
import java.util.Map;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayInputStream;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * <pre>
 * Class Name: ScheduleRecordServiceTest
 * Description: 엑셀 생성 및 파싱 테스트
 *
 * History
 * 2026/05/14 (혜원) 엑셀 생성 및 파싱 테스트 성공
 * </pre>
 */
@SpringBootTest
@Transactional
@DisplayName("스케줄 서비스 통합 테스트")
class ScheduleRecordServiceTest {

    @Autowired private ScheduleRecordService scheduleRecordService;
    @Autowired private ScheduleRecordRepository scheduleRecordRepository;

    // 테스트용 공통 상수
    private static final String YEAR_MONTH  = "2026-05";
    private static final String DEPARTMENT  = "계약 & 교육품질팀";
    private static final String FLIGHT_CODE = "AA, 7L";

    // 각 테스트 실행 전마다 자동으로 실행, 10명의 데이터를 H2 인메모리 DB에 미리 저장하여 테스트
    @BeforeEach
    void setUp() {
        scheduleRecordRepository.saveAll(List.of(
                // 평일 → "9" (출근)
                // 토/일 → "X" (휴무)
                makeRecord(1,  "추순호", may()),
                makeRecord(2,  "문근혁", may("1,15","4,15","5,15","6,15","7,15","8,15","11,15","12,15","13,15","14,15","15,15","18,15","19,15","20,15","21,15","22,15","25,15","26,15","27,15","28,15","29,15")),
                makeRecord(3,  "박태우", may()),
                makeRecord(4,  "김태연", may("5,X","12,X","19,X","26,X")),
                makeRecord(5,  "전극찬", may("5,13","6,13","7,13","12,13","13,13","14,13","19,13","20,13","21,13","26,13","27,13","28,13")),
                makeRecord(6,  "장현지", may()),
                makeRecord(7,  "최다형", may()),
                makeRecord(8,  "서정원", may("1,10","4,10","5,X","6,10","7,4","8,10","11,10","12,X","13,10","14,4","15,10","18,10","19,X","20,10","21,4","22,10","25,10","26,X","27,10","28,4","29,10")),
                makeRecord(9,  "송유영", may("4,11","5,11","6,11","7,11","12,11","13,11","14,11","19,11","20,11","21,11","26,11","27,11","28,11")),
                makeRecord(10, "배종승", may())
        ));
    }

    // ────────────────────────────────────────────────────────────
    // 주 테스트
    // ────────────────────────────────────────────────────────────

    @Test
    @DisplayName("DB에 저장된 스케줄 데이터로 엑셀 파일 생성")
    void generateExcel() throws Exception {

        byte[] excel = scheduleRecordService.generateExcel(
                ScheduleRecordDto.ExcelRequest.builder()
                        .applyYearMonth(YEAR_MONTH)
                        .department(DEPARTMENT)
                        .flightCode(FLIGHT_CODE)
                        .status("확정")
                        .approvers(List.of("부장", "상무"))
                        .build());

        assertThat(excel).isNotNull();
        assertThat(excel.length).isGreaterThan(0);

        // 실제 파일 저장
        Path path = Paths.get("build/test-schedule.xlsx");
        Files.write(path, excel);

        System.out.println("엑셀 저장 완료: " + path.toAbsolutePath());
    }

    @Test
    @DisplayName("생성된 엑셀이 정상적으로 파싱되고 데이터가 일치한다")
    void generateExcelAndParse() throws Exception {

        // given
        byte[] excel = scheduleRecordService.generateExcel(
                ScheduleRecordDto.ExcelRequest.builder()
                        .applyYearMonth(YEAR_MONTH)
                        .department(DEPARTMENT)
                        .flightCode(FLIGHT_CODE)
                        .status("확정")
                        .approvers(List.of("부장", "상무"))
                        .build());

        // then
        assertThat(excel).isNotNull();
        assertThat(excel.length).isGreaterThan(0);

        // 엑셀 파일 저장 (직접 열어보기용)
        Path path = Paths.get("build/test-schedule.xlsx");
        Files.write(path, excel);

        System.out.println("엑셀 저장 완료: " + path.toAbsolutePath());

        // ===== 엑셀 파싱 테스트 =====
        try (Workbook workbook = new XSSFWorkbook(new ByteArrayInputStream(excel))) {

            // 시트 존재 확인
            assertThat(workbook.getNumberOfSheets()).isGreaterThan(0);

            Sheet sheet = workbook.getSheetAt(0);

            System.out.println("시트명: " + sheet.getSheetName());

            // 전체 row 출력 (디버깅용)
            System.out.println("===== 엑셀 내용 =====");

            for (Row row : sheet) {

                StringBuilder sb = new StringBuilder();

                for (Cell cell : row) {

                    String value = switch (cell.getCellType()) {
                        case STRING -> cell.getStringCellValue();
                        case NUMERIC -> String.valueOf((int) cell.getNumericCellValue());
                        case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
                        default -> "";
                    };

                    sb.append(value).append("\t");
                }

                System.out.println(sb);
            }

            // ===== 실제 데이터 검증 =====

            boolean found = false;

            for (Row row : sheet) {

                for (Cell cell : row) {

                    if (cell.getCellType() == CellType.STRING &&
                            cell.getStringCellValue().contains("추순호")) {

                        found = true;

                        System.out.println("추순호 데이터 발견!");
                    }
                }
            }

            assertThat(found).isTrue();
        }
    }

    // ────────────────────────────────────────────────────────────
    // 부 테스트
    // ────────────────────────────────────────────────────────────

    @Test
    @DisplayName("연월 + 부서로 스케줄 조회 시 10명이 순번 순으로 반환된다")
    void getByYearMonthAndDepartment() {
        List<ScheduleRecordDto.Response> result =
                scheduleRecordService.getByYearMonthAndDepartment(YEAR_MONTH, DEPARTMENT);

        assertThat(result).hasSize(10);
        assertThat(result.get(0).getEmployeeName()).isEqualTo("추순호");
        assertThat(result.get(0).getSeq()).isEqualTo(1);
        System.out.println("=== 조회 결과 ===");
        result.forEach(r -> System.out.println(
                r.getSeq() + ". " + r.getEmployeeName() + " | days: " + r.getDays()));
    }

    @Test
    @DisplayName("스케줄 단건 등록 후 id로 조회하면 저장된 days 값이 일치한다")
    void createAndGet() {
        ScheduleRecordDto.Response saved = scheduleRecordService.create(
                ScheduleRecordDto.Request.builder()
                        .applyYearMonth(YEAR_MONTH)
                        .department(DEPARTMENT)
                        .flightCode(FLIGHT_CODE)
                        .employeeName("테스트직원")
                        .seq(11)
                        .days(Map.ofEntries(
                                Map.entry(1,"X"), Map.entry(2,"9"),
                                Map.entry(3,"연"), Map.entry(4,"반"), Map.entry(5,"X/9")))
                        .offDays(10)
                        .usedOff(2.0)
                        .usedAnnual(1.0)
                        .remainAnnual(9.0)
                        .build());

        ScheduleRecordDto.Response found = scheduleRecordService.getById(saved.getId());

        assertThat(found.getEmployeeName()).isEqualTo("테스트직원");
        assertThat(found.getDays().get(3)).isEqualTo("연");
        assertThat(found.getDays().get(5)).isEqualTo("X/9");
        System.out.println("=== 저장된 days ===");
        System.out.println(found.getDays());
    }

    @Test
    @DisplayName("스케줄 수정 시 변경된 날짜값이 정상적으로 반영된다")
    void update() {
        ScheduleRecordDto.Response saved = scheduleRecordService.create(
                ScheduleRecordDto.Request.builder()
                        .applyYearMonth(YEAR_MONTH)
                        .department(DEPARTMENT)
                        .flightCode(FLIGHT_CODE)
                        .employeeName("수정테스트")
                        .seq(12)
                        .days(Map.of(1, "9"))
                        .build());

        ScheduleRecordDto.Response updated = scheduleRecordService.update(
                saved.getId(),
                ScheduleRecordDto.Request.builder()
                        .applyYearMonth(YEAR_MONTH)
                        .department(DEPARTMENT)
                        .flightCode(FLIGHT_CODE)
                        .employeeName("수정테스트")
                        .seq(12)
                        .days(Map.of(1, "X", 2, "연"))
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
                        .applyYearMonth(YEAR_MONTH)
                        .department(DEPARTMENT)
                        .flightCode(FLIGHT_CODE)
                        .employeeName("삭제테스트")
                        .seq(13)
                        .build());

        scheduleRecordService.delete(saved.getId());

        assertThat(scheduleRecordRepository.findById(saved.getId())).isEmpty();
        System.out.println("=== 삭제 완료 ===");
    }

    // ── 헬퍼: 2026년 5월 기본 스케줄 (평일=9, 토일=X)
    // overrides: "날짜,값" 형태로 특정 날짜 덮어쓰기 ex) "5,13"
    @SuppressWarnings("unchecked")
    private static Map<Integer, String> may(String... overrides) {
        String[] base = {
                "9","X","X","9","9","9","9","9","X","X",   // 1~10  (1=금,2=토,3=일)
                "9","9","9","9","9","X","X","9","9","9",    // 11~20
                "9","9","X","X","9","9","9","9","9","X","X" // 21~31
        };
        for (String o : overrides) {
            String[] parts = o.split(",");
            base[Integer.parseInt(parts[0]) - 1] = parts[1];
        }
        Map.Entry<Integer, String>[] entries = new Map.Entry[31];
        for (int i = 0; i < 31; i++) {
            entries[i] = Map.entry(i + 1, base[i]);
        }
        return Map.ofEntries(entries);
    }

    // ── 헬퍼: ScheduleRecord 엔티티 생성
    private ScheduleRecord makeRecord(int seq, String name, Map<Integer, String> days) {
        ScheduleRecord rec = ScheduleRecord.builder()
                .applyYearMonth(YEAR_MONTH)
                .department(DEPARTMENT)
                .flightCode(FLIGHT_CODE)
                .employeeName(name)
                .seq(seq)
                .offDays(10)
                .usedOff(0.0)
                .usedAnnual(0.0)
                .remainAnnual(10.0)
                .build();
        if (days.get(1)  != null) rec.setDay01(days.get(1));
        if (days.get(2)  != null) rec.setDay02(days.get(2));
        if (days.get(3)  != null) rec.setDay03(days.get(3));
        if (days.get(4)  != null) rec.setDay04(days.get(4));
        if (days.get(5)  != null) rec.setDay05(days.get(5));
        if (days.get(6)  != null) rec.setDay06(days.get(6));
        if (days.get(7)  != null) rec.setDay07(days.get(7));
        if (days.get(8)  != null) rec.setDay08(days.get(8));
        if (days.get(9)  != null) rec.setDay09(days.get(9));
        if (days.get(10) != null) rec.setDay10(days.get(10));
        if (days.get(11) != null) rec.setDay11(days.get(11));
        if (days.get(12) != null) rec.setDay12(days.get(12));
        if (days.get(13) != null) rec.setDay13(days.get(13));
        if (days.get(14) != null) rec.setDay14(days.get(14));
        if (days.get(15) != null) rec.setDay15(days.get(15));
        if (days.get(16) != null) rec.setDay16(days.get(16));
        if (days.get(17) != null) rec.setDay17(days.get(17));
        if (days.get(18) != null) rec.setDay18(days.get(18));
        if (days.get(19) != null) rec.setDay19(days.get(19));
        if (days.get(20) != null) rec.setDay20(days.get(20));
        if (days.get(21) != null) rec.setDay21(days.get(21));
        if (days.get(22) != null) rec.setDay22(days.get(22));
        if (days.get(23) != null) rec.setDay23(days.get(23));
        if (days.get(24) != null) rec.setDay24(days.get(24));
        if (days.get(25) != null) rec.setDay25(days.get(25));
        if (days.get(26) != null) rec.setDay26(days.get(26));
        if (days.get(27) != null) rec.setDay27(days.get(27));
        if (days.get(28) != null) rec.setDay28(days.get(28));
        if (days.get(29) != null) rec.setDay29(days.get(29));
        if (days.get(30) != null) rec.setDay30(days.get(30));
        if (days.get(31) != null) rec.setDay31(days.get(31));
        return rec;
    }
}