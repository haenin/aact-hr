package com.aact.overtime.dto;

import lombok.*;
import java.time.LocalDate;
import java.time.LocalTime;

public class OvertimeRecordDto {

    @Getter @Setter @NoArgsConstructor @AllArgsConstructor @Builder
    public static class Request {
        private String department;
        private String employeeName;
        private LocalDate workDate;
        private LocalTime scheduledStart;
        private LocalTime scheduledEnd;
        private LocalTime actualStart;
        private LocalTime actualEnd;
        private Double extensionHours;
        private Double nightHours;
        private Double holidayHours;
        private Double holidayExtensionHours;
        private String applyYearMonth;
    }

    @Getter @Setter @NoArgsConstructor @AllArgsConstructor @Builder
    public static class Response {
        private Long id;
        private String department;
        private String employeeName;
        private LocalDate workDate;
        private LocalTime scheduledStart;
        private LocalTime scheduledEnd;
        private LocalTime actualStart;
        private LocalTime actualEnd;
        private Double extensionHours;
        private Double nightHours;
        private Double holidayHours;
        private Double holidayExtensionHours;
        private String applyYearMonth;
    }

    /** 직원별 소계 */
    @Getter @Setter @NoArgsConstructor @AllArgsConstructor @Builder
    public static class Subtotal {
        private String employeeName;
        private Double extensionHours;      // 연장 소계
        private Double nightHours;          // 야간 소계
        private Double holidayHours;        // 휴일 소계
        private Double holidayExtensionHours; // 휴일연장 소계
    }

    /** 부서 전체 합계 */
    @Getter @Setter @NoArgsConstructor @AllArgsConstructor @Builder
    public static class Total {
        private String department;
        private Double extensionHours;      // 연장 합계
        private Double nightHours;          // 야간 합계
        private Double holidayHours;        // 휴일 합계
        private Double holidayExtensionHours; // 휴일연장 합계
    }
}