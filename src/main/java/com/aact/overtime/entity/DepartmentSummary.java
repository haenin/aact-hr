package com.aact.overtime.entity;

import jakarta.persistence.*;
import lombok.*;

/**
 * 부서별 시간외근무 집계 (표지 집계표)
 */
@Entity
@Table(name = "overtime_department_summary",
       uniqueConstraints = @UniqueConstraint(columnNames = {"department", "applyYearMonth"}))
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class DepartmentSummary {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    @Column(nullable = false, length = 50)
    private String department;           // 부서명

    @Column(nullable = false, length = 7)
    private String applyYearMonth;       // 신청 연월 (ex: "2026-04")

    // 전월 실적
    private Double prevExtensionHours;
    private Double prevNightHours;
    private Double prevHolidayHours;

    // 당월 실적
    private Double currExtensionHours;
    private Double currNightHours;
    private Double currHolidayHours;

    // 증감 (당월 - 전월)
    private Double diffExtensionHours;
    private Double diffNightHours;
    private Double diffHolidayHours;

    // 증감 사유 (OT사유)
    @Column(length = 500)
    private String changeReason;
}
