package com.aact.overtime.dto;

import lombok.*;

public class DepartmentSummaryDto {

    @Getter
    @Setter
    @NoArgsConstructor
    @AllArgsConstructor
    @Builder
    public static class Request {
        private String department;
        private String applyYearMonth;
        private Double prevExtensionHours;
        private Double prevNightHours;
        private Double prevHolidayHours;
        private Double currExtensionHours;
        private Double currNightHours;
        private Double currHolidayHours;
        private String changeReason;
    }

    @Getter
    @Setter
    @NoArgsConstructor
    @AllArgsConstructor
    @Builder
    public static class Response {
        private Long id;
        private String department;
        private String applyYearMonth;
        private Double prevExtensionHours;
        private Double prevNightHours;
        private Double prevHolidayHours;
        private Double currExtensionHours;
        private Double currNightHours;
        private Double currHolidayHours;
        private Double diffExtensionHours;
        private Double diffNightHours;
        private Double diffHolidayHours;
        private String changeReason;
    }

    /** 부서별 전체 시간 합산 (개인별 집계 롤업용) */
    @Getter
    @Setter
    @NoArgsConstructor
    @AllArgsConstructor
    @Builder
    public static class AggregateResult {
        private String department;
        private Double totalExtension;
        private Double totalNight;
        private Double totalHoliday;
        private Double totalHolidayExtension;
    }
}
