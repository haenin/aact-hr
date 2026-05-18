package com.aact.schedule.fixture;

import com.aact.schedule.dto.ScheduleRecordDto;
import com.aact.schedule.entity.ScheduleRecord;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * <pre>
 * Class Name: ScheduleTestFixture
 * Description: 스케줄 테스트 데이터 픽스처
 *
 *   Dataset A — 정상 케이스  (계약 & 교육품질팀 · AA, 7L · 10인)
 *               평일 9 / 토일 X 기본패턴 + 교대·연차·특수코드 혼합
 *               결재: 대리 · 부장 · 상무
 *
 *   Dataset B — 교대근무 케이스 (화물운영팀 · KE, OZ · 6인)
 *               오전/오후/야간 교대 + 반차·대체휴무 혼재
 *               결재: 사원 · 대리 · 팀장
 *
 *   Dataset C — 스트레스 케이스 (수출입지원팀 · AA, CX · 15인)
 *               연차·반차·대체·X/9 복합코드 총망라
 *               결재: 주임 · 과장 · 차장 · 상무
 *
 * History
 * 2026/05/15 (혜원) 테스트 픽스처 분리 작성
 * </pre>
 */
public final class ScheduleTestFixture {

    // ── 공통
    public static final String YEAR_MONTH = "2026-05";

    // ── 부서 / 편명
    public static final String DEPT_A        = "계약 & 교육품질팀";
    public static final String DEPT_B        = "화물운영팀";
    public static final String DEPT_C        = "수출입지원팀";
    public static final String FLIGHT_CODE_A = "AA, 7L";
    public static final String FLIGHT_CODE_B = "KE, OZ";
    public static final String FLIGHT_CODE_C = "AA, CX";

    // ── 결재자 (직급명)
    public static final List<String> APPROVERS_A = List.of("부장", "상무");
    public static final List<String> APPROVERS_B = List.of("대리", "팀장");
    public static final List<String> APPROVERS_C = List.of("주임", "과장", "차장", "상무");

    private ScheduleTestFixture() {}

    // ────────────────────────────────────────────────────────────────
    // Dataset A — 정상 케이스 (10인)
    // ────────────────────────────────────────────────────────────────
    public static List<ScheduleRecord> datasetA() {
        return List.of(
                make(1,  DEPT_A, FLIGHT_CODE_A, "추순호", may()),
                make(2,  DEPT_A, FLIGHT_CODE_A, "문근혁", may(
                        "1,15","4,15","5,15","6,15","7,15","8,15",
                        "11,15","12,15","13,15","14,15","15,15",
                        "18,15","19,15","20,15","21,15","22,15",
                        "25,15","26,15","27,15","28,15","29,15")),
                make(3,  DEPT_A, FLIGHT_CODE_A, "박태우", may()),
                make(4,  DEPT_A, FLIGHT_CODE_A, "김태연", may("5,X","12,X","19,X","26,X")),
                make(5,  DEPT_A, FLIGHT_CODE_A, "전극찬", may(
                        "5,13","6,13","7,13","12,13","13,13","14,13",
                        "19,13","20,13","21,13","26,13","27,13","28,13")),
                make(6,  DEPT_A, FLIGHT_CODE_A, "장현지", may()),
                make(7,  DEPT_A, FLIGHT_CODE_A, "최다형", may()),
                make(8,  DEPT_A, FLIGHT_CODE_A, "서정원", may(
                        "1,10","4,10","5,X","6,10","7,4","8,10",
                        "11,10","12,X","13,10","14,4","15,10",
                        "18,10","19,X","20,10","21,4","22,10",
                        "25,10","26,X","27,10","28,4","29,10")),
                make(9,  DEPT_A, FLIGHT_CODE_A, "송유영", may(
                        "4,11","5,11","6,11","7,11","12,11","13,11","14,11",
                        "19,11","20,11","21,11","26,11","27,11","28,11")),
                make(10, DEPT_A, FLIGHT_CODE_A, "배종승", may())
        );
    }

    // ────────────────────────────────────────────────────────────────
    // Dataset B — 교대근무 케이스 (6인)
    // ────────────────────────────────────────────────────────────────
    public static List<ScheduleRecord> datasetB() {
        return List.of(
                // 오전(4시) 교대
                make(1, DEPT_B, FLIGHT_CODE_B, "강민준", may(
                        "1,4","4,4","5,4","6,4","7,4","8,4",
                        "11,4","12,4","13,4","14,4","15,4",
                        "18,4","19,4","20,4","21,4","22,4",
                        "25,4","26,4","27,4","28,4","29,4")),
                // 오후(13시) 교대
                make(2, DEPT_B, FLIGHT_CODE_B, "윤서아", may(
                        "1,13","4,13","5,13","6,13","7,13","8,13",
                        "11,13","12,13","13,13","14,13","15,13",
                        "18,13","19,13","20,13","21,13","22,13",
                        "25,13","26,13","27,13","28,13","29,13")),
                // 야간(22시) 교대 + 연차 2일
                make(3, DEPT_B, FLIGHT_CODE_B, "박도현", may(
                        "1,22","4,22","5,22","6,22","7,22","8,22",
                        "11,연","12,연",
                        "13,22","14,22","15,22",
                        "18,22","19,22","20,22","21,22","22,22",
                        "25,22","26,22","27,22","28,22","29,22")),
                // 오전/야간 격주 + 대체휴무
                make(4, DEPT_B, FLIGHT_CODE_B, "최예은", may(
                        "1,4","4,4","5,4","6,4","7,4","8,4",
                        "11,대","12,대",
                        "13,22","14,22","15,22",
                        "18,22","19,22","20,22","21,22","22,22",
                        "25,4","26,4","27,4","28,4","29,4")),
                // 오후 교대 + 반차 격주
                make(5, DEPT_B, FLIGHT_CODE_B, "임지호", may(
                        "1,13","4,13","5,반","6,13","7,13","8,13",
                        "11,13","12,반","13,13","14,13","15,13",
                        "18,13","19,반","20,13","21,13","22,13",
                        "25,13","26,반","27,13","28,13","29,13")),
                // 오전(10시) 교대 + 토·일 근무 포함
                make(6, DEPT_B, FLIGHT_CODE_B, "오태양", may(
                        "1,10","2,10","4,10","5,10","6,10","7,10","8,10","9,10",
                        "11,10","12,10","13,10","14,10","15,10",
                        "18,10","19,10","20,10","21,10","22,10",
                        "25,10","26,10","27,10","28,10","29,10"))
        );
    }

    // ────────────────────────────────────────────────────────────────
    // Dataset C — 스트레스 케이스 (15인)
    // ────────────────────────────────────────────────────────────────
    public static List<ScheduleRecord> datasetC() {
        Object[][] rows = {
                {1,  "강하은", new String[]{}},
                {2,  "권민서", new String[]{"5,연","6,연","7,연"}},
                {3,  "김도윤", new String[]{"12,반","13,반","19,반","26,반"}},
                {4,  "남지우", new String[]{"1,4","4,4","5,4","8,4","11,4","12,4","13,4","15,4","18,4","19,4","20,4","22,4","25,4","26,4","27,4","29,4"}},
                {5,  "노재원", new String[]{"4,13","5,13","6,13","7,13","8,13","11,13","12,13","13,13","14,13","15,13","18,13","19,13","20,13","21,13","22,13","25,13","26,13","27,13","28,13","29,13"}},
                {6,  "류승현", new String[]{"5,X/9","12,X/9","19,X/9","26,X/9"}},
                {7,  "문지아", new String[]{"1,대","4,대","11,연","12,연","13,연","18,대","19,대"}},
                {8,  "박서윤", new String[]{}},
                {9,  "서하준", new String[]{"4,10","5,10","6,10","7,10","8,10","11,10","12,10","13,10","14,10","15,10","18,10","19,10","20,10","21,10","22,10","25,10","26,10","27,10","28,10","29,10"}},
                {10, "송이현", new String[]{"21,연","22,연","25,연","26,연","27,연","28,연","29,연"}},
                {11, "신민재", new String[]{"1,22","4,22","5,22","6,22","7,22","8,22","11,22","12,22","13,22","14,22","15,22","18,22","19,22","20,22","21,22","22,22","25,22","26,22","27,22","28,22","29,22"}},
                {12, "안지수", new String[]{"5,반","7,반","12,반","14,반","19,반","21,반","26,반","28,반"}},
                {13, "이하율", new String[]{}},
                {14, "정수아", new String[]{"4,15","5,15","6,15","7,15","8,15","11,15","12,15","13,15","14,15","15,15","18,15","19,15","20,15","21,15","22,15","25,15","26,15","27,15","28,15","29,15"}},
                {15, "최다온", new String[]{"1,9","4,9","5,대","6,9","7,대","11,9","12,대","13,9","14,대","15,9","18,9","19,대","20,9","21,대","22,9","25,9","26,대","27,9","28,대","29,9"}},
        };

        List<ScheduleRecord> list = new ArrayList<>();
        for (Object[] row : rows) {
            list.add(make((int) row[0], DEPT_C, FLIGHT_CODE_C, (String) row[1], may((String[]) row[2])));
        }
        return list;
    }

    // ────────────────────────────────────────────────────────────────
    // ExcelRequest 빌더
    // ────────────────────────────────────────────────────────────────
    public static ScheduleRecordDto.ExcelRequest requestFor(
            String department, String flightCode, List<String> approvers) {
        return ScheduleRecordDto.ExcelRequest.builder()
                .applyYearMonth(YEAR_MONTH)
                .department(department)
                .flightCode(flightCode)
                .status("확정")
                .approvers(approvers)
                .build();
    }

    // ────────────────────────────────────────────────────────────────
    // 헬퍼 — 2026년 5월 기본 패턴 (평일=9, 토일=X)
    // overrides: "날짜,값" 형태. ex) "5,13", "12,연"
    // ────────────────────────────────────────────────────────────────
    @SuppressWarnings("unchecked")
    public static Map<Integer, String> may(String... overrides) {
        String[] base = {
                "9","X","X","9","9","9","9","9","X","X",   // 1~10  (1=금,2=토,3=일)
                "9","9","9","9","9","X","X","9","9","9",    // 11~20
                "9","9","X","X","9","9","9","9","9","X","X" // 21~31
        };
        for (String o : overrides) {
            String[] parts = o.split(",", 2);
            base[Integer.parseInt(parts[0]) - 1] = parts[1];
        }
        Map.Entry<Integer, String>[] entries = new Map.Entry[31];
        for (int i = 0; i < 31; i++) entries[i] = Map.entry(i + 1, base[i]);
        return Map.ofEntries(entries);
    }

    // ────────────────────────────────────────────────────────────────
    // 헬퍼 — ScheduleRecord 엔티티 생성
    // ────────────────────────────────────────────────────────────────
    private static ScheduleRecord make(int seq, String department, String flightCode,
                                       String name, Map<Integer, String> days) {
        ScheduleRecord rec = ScheduleRecord.builder()
                .applyYearMonth(YEAR_MONTH)
                .department(department)
                .flightCode(flightCode)
                .employeeName(name)
                .seq(seq)
                .offDays(10)
                .usedOff(0.0)
                .usedAnnual(0.0)
                .remainAnnual(10.0)
                .build();

        for (int d = 1; d <= 31; d++) {
            String val = days.get(d);
            if (val == null) continue;
            switch (d) {
                case  1 -> rec.setDay01(val); case  2 -> rec.setDay02(val);
                case  3 -> rec.setDay03(val); case  4 -> rec.setDay04(val);
                case  5 -> rec.setDay05(val); case  6 -> rec.setDay06(val);
                case  7 -> rec.setDay07(val); case  8 -> rec.setDay08(val);
                case  9 -> rec.setDay09(val); case 10 -> rec.setDay10(val);
                case 11 -> rec.setDay11(val); case 12 -> rec.setDay12(val);
                case 13 -> rec.setDay13(val); case 14 -> rec.setDay14(val);
                case 15 -> rec.setDay15(val); case 16 -> rec.setDay16(val);
                case 17 -> rec.setDay17(val); case 18 -> rec.setDay18(val);
                case 19 -> rec.setDay19(val); case 20 -> rec.setDay20(val);
                case 21 -> rec.setDay21(val); case 22 -> rec.setDay22(val);
                case 23 -> rec.setDay23(val); case 24 -> rec.setDay24(val);
                case 25 -> rec.setDay25(val); case 26 -> rec.setDay26(val);
                case 27 -> rec.setDay27(val); case 28 -> rec.setDay28(val);
                case 29 -> rec.setDay29(val); case 30 -> rec.setDay30(val);
                case 31 -> rec.setDay31(val);
            }
        }
        return rec;
    }
}
