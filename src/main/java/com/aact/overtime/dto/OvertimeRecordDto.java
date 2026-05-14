package com.aact.overtime.dto;

import lombok.*;
import java.time.LocalDate;
import java.time.LocalTime;

public class OvertimeRecordDto {

    @Getter
    @Setter
    @NoArgsConstructor
    @AllArgsConstructor
    @Builder
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
        private String applyYearMonth;   // "2026-04"
    }

    @Getter
    @Setter
    @NoArgsConstructor
    @AllArgsConstructor
    @Builder
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
}
