package com.aact.overtime.service;

import com.aact.overtime.dto.ApplicationDto;
import com.aact.overtime.repository.ApplicationRepository;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.transaction.annotation.Transactional;

import java.io.ByteArrayInputStream;
import java.util.List;

import static com.aact.overtime.fixture.ApplicationTestFixture.*;
import static org.assertj.core.api.Assertions.assertThat;

@SpringBootTest
@Transactional
@DisplayName("시간외근무수당 신청서 엑셀 파싱 테스트")
class ApplicationExcelParserTest {

    @Autowired
    private ApplicationService applicationService;
    @Autowired
    private ApplicationExcelParser applicationExcelParser;
    @Autowired
    private ApplicationRepository applicationRepository;

    @BeforeEach
    void setUp() {
        applicationRepository.deleteAll();
    }

    @Test
    @DisplayName("생성된 엑셀을 파싱하면 신청서 데이터가 정상 출력된다")
    void parse_generatedExcel_printAll() throws Exception {
        // given
        applicationRepository.saveAll(datasetA());
        byte[] excel = applicationService.generateExcel(requestFor(DEPT_A));

        // when
        List<ApplicationDto.Request> parsed =
                applicationExcelParser.parse(
                        new ByteArrayInputStream(excel), YEAR_MONTH, DEPT_A);

        List<String> approvers =
                applicationExcelParser.parseApprovers(new ByteArrayInputStream(excel));

        // then
        assertThat(parsed).isNotEmpty();

        // ── 결재란 출력
        System.out.println("\n");
        System.out.println("╔══════════════════════════════════════════════════════════════════════════════╗");
        System.out.printf("║          [신청서 파싱 결과] %-50s║%n", DEPT_A + " · " + YEAR_MONTH);
        System.out.println("╠══════════════════════════════════════════════════════════════════════════════╣");
        System.out.println("║ [결재란]                                                                      ║");
        if (approvers.isEmpty()) {
            System.out.println("║   (없음)                                                                     ║");
        } else {
            approvers.forEach(a ->
                    System.out.printf("║   - %-73s║%n", a.replace("\n", " / "))
            );
        }

        // ── 데이터 출력
        System.out.println("╠════╦══════════╦══════╦══════════════╦══════════════╦══════╦══════╦══════╦══════════════════════╣");
        System.out.println("║ NO ║   일자   ║ 성명 ║  예정근무시간  ║  실근무시간   ║ 연장 ║ 야간 ║ 휴일 ║       근무사유        ║");
        System.out.println("╠════╬══════════╬══════╬══════════════╬══════════════╬══════╬══════╬══════╬══════════════════════╣");

        for (ApplicationDto.Request r : parsed) {
            System.out.printf("║ %2d ║ %-8s ║ %-4s ║ %-12s ║ %-12s ║ %4.1f ║ %4.1f ║ %4.1f ║ %-20s ║%n",
                    r.getNo(),
                    nvl(r.getWorkDate()),
                    nvl(r.getEmployeeName()),
                    nvl(r.getScheduledTime()),
                    nvl(r.getActualTime()),
                    r.getExtensionHours() != null ? r.getExtensionHours() : 0.0,
                    r.getNightHours() != null ? r.getNightHours() : 0.0,
                    r.getHolidayHours() != null ? r.getHolidayHours() : 0.0,
                    nvl(r.getWorkReason()));
        }

        System.out.println("╚════╩══════════╩══════╩══════════════╩══════════════╩══════╩══════╩══════╩══════════════════════╝");
        System.out.println("  → 총 " + parsed.size() + "건 파싱 / 결재란 " + approvers.size() + "개 파싱");
    }

    private String nvl(String s) {
        return s != null ? s : "-";
    }
}