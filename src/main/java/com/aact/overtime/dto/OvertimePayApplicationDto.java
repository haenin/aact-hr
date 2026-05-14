package com.aact.overtime.dto;

import lombok.*;
import java.util.List;

public class OvertimePayApplicationDto {

    @Getter @Setter @NoArgsConstructor @AllArgsConstructor @Builder
    public static class Request {
        private String applyYearMonth;  // "2026-03"
        private String department;      // 부서명
        private Integer no;             // 순번
        private String workDate;        // 일자
        private String employeeName;    // 성명
        private String scheduledTime;   // 예정근무시간
        private String actualTime;      // 실 근무시간
        private Double extensionHours;  // 연장
        private Double nightHours;      // 야간
        private Double holidayHours;    // 휴일
        private String workReason;      // 근무사유
    }

    @Getter @Setter @NoArgsConstructor @AllArgsConstructor @Builder
    public static class Response {
        private Long id;
        private String applyYearMonth;
        private String department;
        private Integer no;
        private String workDate;
        private String employeeName;
        private String scheduledTime;
        private String actualTime;
        private Double extensionHours;
        private Double nightHours;
        private Double holidayHours;
        private String workReason;
    }

    /** 서명자 (신청자/팀장/부서장 등 동적) */
    @Getter @Setter @NoArgsConstructor @AllArgsConstructor @Builder
    public static class Signer {
        private String role;    // ex) "신청자", "팀 장", "부서장" (동적)
        private String name;    // ex) "서정원"
    }

    /** 엑셀 생성 요청 */
    @Getter @Setter @NoArgsConstructor @AllArgsConstructor @Builder
    public static class ExcelRequest {
        private String applyYearMonth;      // "2026-03"
        private String department;          // 부서명
        private List<String> approvers;     // 결재란 직급 ex) ["담당", "대리"] (동적)
        private List<Signer> signers;       // 서명란 ex) [{신청자, 서정원}, {팀 장, 배종승}, {부서장, 김해균}]
    }
}
