package com.aact.schedule;

import com.aact.schedule.dto.ScheduleRecordDto;
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
@DisplayName("스케줄 엑셀 파싱 테스트")
class ScheduleExcelParserTest {

    @Autowired private ScheduleExcelParser scheduleExcelParser;

    private static final String YEAR_MONTH = "2026-05";
    private static final String DEPT       = "항공기운송지원팀";

    @Test
    @DisplayName("실제 엑셀 파일을 파싱하면 스케줄 데이터가 정상 출력된다")
    void parse_realExcel_printAll() throws Exception {
        try (InputStream is = getClass().getResourceAsStream("/fixture/test-schedule-A.xlsx")) {

            assertThat(is).as("fixture/test-schedule-A.xlsx 파일이 test/resources에 없습니다").isNotNull();
            byte[] bytes = is.readAllBytes();

            // when
            List<ScheduleRecordDto.ParsedRow> parsed =
                    scheduleExcelParser.parse(new ByteArrayInputStream(bytes), YEAR_MONTH, DEPT);

            List<String> approvers =
                    scheduleExcelParser.parseApprovers(new ByteArrayInputStream(bytes));

            // then
            assertThat(parsed).isNotEmpty();

            // ── 결재란 출력
            System.out.println("\n");
            System.out.println("╔══════════════════════════════════════════════════════════════════╗");
            System.out.printf( "║        [스케줄 파싱 결과] %-40s║%n", DEPT + " · " + YEAR_MONTH);
            System.out.println("╠══════════════════════════════════════════════════════════════════╣");
            System.out.println("║ [결재란]                                                          ║");
            if (approvers.isEmpty()) {
                System.out.println("║   (없음)                                                         ║");
            } else {
                approvers.forEach(a ->
                        System.out.printf("║   - %-62s║%n", a.replace("\n", " / "))
                );
            }

            // ── 데이터 출력
            System.out.println("╠════╦══════╦══════════════════════════════════════════╦════╦════╦════╦════╣");
            System.out.println("║ NO ║ 성명 ║               1일~31일 근무패턴             ║휴무║사휴║사연║잔연║");
            System.out.println("╠════╬══════╬══════════════════════════════════════════╬════╬════╬════╬════╣");

            for (ScheduleRecordDto.ParsedRow row : parsed) {
                String dayPattern = String.join("|", row.getDays());
                if (dayPattern.length() > 42) dayPattern = dayPattern.substring(0, 39) + "...";

                System.out.printf("║ %2d ║ %-4s ║ %-42s ║ %-2s ║ %-2s ║ %-2s ║ %-2s ║%n",
                        row.getSeq(),
                        row.getEmployeeName(),
                        dayPattern,
                        nvl(row.getOffDays()),
                        nvl(row.getUsedOff()),
                        nvl(row.getUsedAnnual()),
                        nvl(row.getRemainAnnual()));
            }

            System.out.println("╚════╩══════╩══════════════════════════════════════════╩════╩════╩════╩════╝");
            System.out.println("  → 총 " + parsed.size() + "건 파싱 / 결재란 " + approvers.size() + "개 파싱");
        }
    }

    private String nvl(String s) {
        return s != null ? s : "-";
    }
}