package com.aact.schedule.dto;

import lombok.*;
import java.util.List;
import java.util.Map;

public class ScheduleRecordDto {

    @Getter
    @Setter
    @NoArgsConstructor
    @AllArgsConstructor
    @Builder
    public static class Request {
        private String applyYearMonth;     // "2026-05"
        private String department;         // 부서명
        private String flightCode;         // 편명
        private String status;             // 확정 상태 ex) "확정", "임시"
        private String employeeName;       // 성명 ex) "추순호", "문근혁"
        private Integer seq;               // 순번
        private Map<Integer, String> days; // key: 1~31, value: "X","9" 등
        private Integer offDays;           // 휴무일수
        private Double usedOff;            // 사용휴무
        private Double usedAnnual;         // 사용연차
        private Double remainAnnual;       // 잔여연차
    }

    @Getter
    @Setter
    @NoArgsConstructor
    @AllArgsConstructor
    @Builder
    public static class Response {
        private Long id;
        private String applyYearMonth;
        private String department;
        private String flightCode;
        private String employeeName;
        private Integer seq;
        private Map<Integer, String> days; // key: 1~31
        private Integer offDays;
        private Double usedOff;
        private Double usedAnnual;
        private Double remainAnnual;
    }

    /** 엑셀 생성 요청 */
    @Getter
    @Setter
    @NoArgsConstructor
    @AllArgsConstructor
    @Builder
    public static class ExcelRequest {
        private String applyYearMonth;     // "2026-05"
        private String department;         // 부서명
        private String flightCode;         // 편명
        private String status;             // 확정 상태 ex) "확정", "임시"
        private List<String> approvers;    // 결재자 목록 ex) ["부장", "상무"] (동적)
    }
}