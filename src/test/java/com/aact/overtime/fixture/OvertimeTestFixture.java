package com.aact.overtime.fixture;

import com.aact.overtime.dto.DepartmentSummaryDto;
import com.aact.overtime.entity.DepartmentSummary;
import com.aact.overtime.entity.OvertimeRecord;

import java.time.LocalDate;
import java.time.LocalTime;
import java.util.List;

/**
 * <pre>
 * Class Name: OvertimeTestFixture
 * Description: 시간외근무 엑셀 테스트 데이터 픽스처
 *
 *   Dataset A — 정상 케이스 (항공기운송지원팀 · 3인 7건)
 *               연장/야간 혼합 · 김보라 연장 소계 9.5h
 *
 *   Dataset B — 경계값 케이스 (화물영업팀 · 야간 최대 + 휴일 포함)
 *               야간 8h 최대치, 휴일근무 포함
 *
 *   Dataset C — 스트레스 케이스 (3개 부서 집계 + 대량 세부내역)
 *               부서별 집계표 3건 + 세부내역 다수
 *
 * History
 * 2026/05/18 (혜원) 테스트 픽스처 분리 작성
 * </pre>
 */
public final class OvertimeTestFixture {

    public static final String YEAR_MONTH = "2026-05";
    public static final String DEPT_A = "항공기운송지원팀";
    public static final String DEPT_B = "화물영업팀";
    public static final String DEPT_C = "수출입통관팀";

    public static final List<String> APPROVERS_4 = List.of("담당", "과장", "상무", "부사장");
    public static final List<String> APPROVERS_2 = List.of("담당", "부장");

    private OvertimeTestFixture() {}

    // ────────────────────────────────────────────────────────────────
    // Dataset A — 정상 케이스
    // 항공기운송지원팀 · 3인 7건 · 연장/야간 혼합
    // 김보라 연장 소계: 2.5 + 3.0 + 4.0 = 9.5
    // 부서 연장 합계: 18.0 / 야간: 9.0 / 휴일: 0.0
    // ────────────────────────────────────────────────────────────────
    public static List<OvertimeRecord> recordsA() {
        return List.of(
                makeRecord(DEPT_A, "김보라", LocalDate.of(2026, 5, 1),
                        "09:00", "18:00", "09:00", "20:30", 2.5, 0.0, 0.0, 0.0),
                makeRecord(DEPT_A, "김보라", LocalDate.of(2026, 5, 8),
                        "09:00", "18:00", "09:00", "21:00", 3.0, 0.0, 0.0, 0.0),
                makeRecord(DEPT_A, "김보라", LocalDate.of(2026, 5, 15),
                        "09:00", "18:00", "09:00", "22:00", 4.0, 1.0, 0.0, 0.0),
                makeRecord(DEPT_A, "모은서", LocalDate.of(2026, 5, 2),
                        "09:00", "18:00", "09:00", "20:00", 2.0, 0.0, 0.0, 0.0),
                makeRecord(DEPT_A, "모은서", LocalDate.of(2026, 5, 9),
                        "09:00", "18:00", "09:00", "23:00", 5.0, 0.0, 0.0, 0.0),
                makeRecord(DEPT_A, "남다정", LocalDate.of(2026, 5, 3),
                        "13:00", "22:00", "13:00", "22:00", 0.0, 4.0, 0.0, 0.0),
                makeRecord(DEPT_A, "남다정", LocalDate.of(2026, 5, 10),
                        "13:00", "22:00", "13:00", "23:30", 1.5, 4.0, 0.0, 0.0)
        );
    }

    public static List<DepartmentSummary> summariesA() {
        return List.of(
                makeSummary(DEPT_A, 12.0, 3.0, 0.0, 18.0, 9.0, 0.0, "항공편 증가")
        );
    }

    // ────────────────────────────────────────────────────────────────
    // Dataset B — 경계값 케이스
    // 화물영업팀 · 4인 8건 · 야간 최대치(8h) + 휴일 포함
    // 부서 연장 합계: 7.5 / 야간: 18.0 / 휴일: 16.0
    // ────────────────────────────────────────────────────────────────
    public static List<OvertimeRecord> recordsB() {
        return List.of(
                makeRecord(DEPT_B, "이준혁", LocalDate.of(2026, 5, 2),
                        "09:00", "18:00", "09:00", "21:00", 3.0, 0.0, 0.0, 0.0),
                makeRecord(DEPT_B, "이준혁", LocalDate.of(2026, 5, 9),
                        "09:00", "18:00", "09:00", "19:30", 1.5, 0.0, 0.0, 0.0),
                makeRecord(DEPT_B, "정수빈", LocalDate.of(2026, 5, 6),
                        "22:00", "06:00", "22:00", "06:00", 0.0, 8.0, 0.0, 0.0),
                makeRecord(DEPT_B, "정수빈", LocalDate.of(2026, 5, 13),
                        "22:00", "06:00", "22:00", "06:00", 0.0, 8.0, 0.0, 0.0),
                makeRecord(DEPT_B, "한소희", LocalDate.of(2026, 5, 5),
                        "00:00", "00:00", "09:00", "17:00", 0.0, 0.0, 8.0, 0.0),
                makeRecord(DEPT_B, "한소희", LocalDate.of(2026, 5, 25),
                        "00:00", "00:00", "09:00", "17:00", 0.0, 0.0, 8.0, 0.0),
                makeRecord(DEPT_B, "오태양", LocalDate.of(2026, 5, 20),
                        "15:00", "24:00", "15:00", "24:00", 3.0, 2.0, 0.0, 0.0),
                makeRecord(DEPT_B, "오태양", LocalDate.of(2026, 5, 22),
                        "15:00", "24:00", "15:00", "24:00", 0.0, 0.0, 0.0, 0.0)
        );
    }

    public static List<DepartmentSummary> summariesB() {
        return List.of(
                makeSummary(DEPT_B, 5.0, 12.0, 8.0, 7.5, 18.0, 16.0, "야간 운항 증가")
        );
    }

    // ────────────────────────────────────────────────────────────────
    // Dataset C — 스트레스 케이스
    // 3개 부서 집계표 + 항공기운송지원팀 세부내역 다수
    // ────────────────────────────────────────────────────────────────
    public static List<DepartmentSummary> summariesC() {
        return List.of(
                makeSummary("조업부",    20.0, 5.0, 3.0, 22.0, 6.0, 2.0, "업무증가"),
                makeSummary(DEPT_A,     15.0, 3.0, 2.0, 18.0, 4.0, 3.0, "항공편 증가"),
                makeSummary("관리부",   10.0, 2.0, 1.0, 12.0, 2.0, 1.0, null)
        );
    }

    public static List<OvertimeRecord> recordsC() {
        // recordsA + recordsB 합산으로 대량 데이터 구성
        return List.of(
                makeRecord(DEPT_A, "김보라", LocalDate.of(2026, 5, 1),
                        "09:00", "18:00", "09:00", "20:30", 2.5, 0.0, 0.0, 0.0),
                makeRecord(DEPT_A, "김보라", LocalDate.of(2026, 5, 8),
                        "09:00", "18:00", "09:00", "21:00", 3.0, 0.0, 0.0, 0.0),
                makeRecord(DEPT_A, "김보라", LocalDate.of(2026, 5, 15),
                        "09:00", "18:00", "09:00", "22:00", 4.0, 1.0, 0.0, 0.0),
                makeRecord(DEPT_A, "모은서", LocalDate.of(2026, 5, 2),
                        "09:00", "18:00", "09:00", "20:00", 2.0, 0.0, 0.0, 0.0),
                makeRecord(DEPT_A, "모은서", LocalDate.of(2026, 5, 9),
                        "09:00", "18:00", "09:00", "23:00", 5.0, 0.0, 0.0, 0.0),
                makeRecord(DEPT_A, "남다정", LocalDate.of(2026, 5, 3),
                        "13:00", "22:00", "13:00", "22:00", 0.0, 4.0, 0.0, 0.0),
                makeRecord(DEPT_A, "남다정", LocalDate.of(2026, 5, 10),
                        "13:00", "22:00", "13:00", "23:30", 1.5, 4.0, 0.0, 0.0),
                makeRecord(DEPT_A, "이준혁", LocalDate.of(2026, 5, 2),
                        "09:00", "18:00", "09:00", "21:00", 3.0, 0.0, 0.0, 0.0),
                makeRecord(DEPT_A, "이준혁", LocalDate.of(2026, 5, 9),
                        "09:00", "18:00", "09:00", "19:30", 1.5, 0.0, 0.0, 0.0),
                makeRecord(DEPT_A, "정수빈", LocalDate.of(2026, 5, 6),
                        "22:00", "06:00", "22:00", "06:00", 0.0, 8.0, 0.0, 0.0),
                makeRecord(DEPT_A, "정수빈", LocalDate.of(2026, 5, 13),
                        "22:00", "06:00", "22:00", "06:00", 0.0, 8.0, 0.0, 0.0),
                makeRecord(DEPT_A, "한소희", LocalDate.of(2026, 5, 5),
                        "00:00", "00:00", "09:00", "17:00", 0.0, 0.0, 8.0, 0.0),
                makeRecord(DEPT_A, "한소희", LocalDate.of(2026, 5, 25),
                        "00:00", "00:00", "09:00", "17:00", 0.0, 0.0, 8.0, 0.0)
        );
    }

    // ────────────────────────────────────────────────────────────────
    // 내부 헬퍼
    // ────────────────────────────────────────────────────────────────
    public static OvertimeRecord makeRecord(
            String dept, String name, LocalDate date,
            String schStart, String schEnd,
            String actStart, String actEnd,
            double ext, double night, double hol, double holExt) {
        return OvertimeRecord.builder()
                .applyYearMonth(YEAR_MONTH)
                .department(dept)
                .employeeName(name)
                .workDate(date)
                .scheduledStart(parseTime(schStart))
                .scheduledEnd(parseTime(schEnd))
                .actualStart(parseTime(actStart))
                .actualEnd(parseTime(actEnd))
                .extensionHours(ext)
                .nightHours(night)
                .holidayHours(hol)
                .holidayExtensionHours(holExt)
                .build();
    }

    // 24:00 → 00:00 처리
    private static LocalTime parseTime(String time) {
        if ("24:00".equals(time)) return LocalTime.MIDNIGHT;
        return LocalTime.parse(time);
    }

    public static DepartmentSummary makeSummary(
            String dept,
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
}