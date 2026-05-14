package com.aact.overtime.entity;

import jakarta.persistence.*;
import lombok.*;
import java.time.LocalDate;
import java.time.LocalTime;

/**
 * 시간외근무 개인별 세부내역
 */
@Entity
@Table(name = "overtime_record")
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class OvertimeRecord {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    // 직원 정보
    @Column(nullable = false, length = 50)
    private String department;          // 부서명

    @Column(nullable = false, length = 50)
    private String employeeName;        // 성명

    @Column(nullable = false)
    private LocalDate workDate;         // 근무 일자

    // 예정 근무
    private LocalTime scheduledStart;   // 예정 출근
    private LocalTime scheduledEnd;     // 예정 퇴근

    // 실근무 (지문)
    private LocalTime actualStart;      // 실근무 출근
    private LocalTime actualEnd;        // 실근무 퇴근

    // 시간외 근무 시간 (단위: 시간, 소수점 포함 ex: 1.5)
    @Column(nullable = false)
    private Double extensionHours;      // 연장근무

    @Column(nullable = false)
    private Double nightHours;          // 야간근무

    @Column(nullable = false)
    private Double holidayHours;        // 휴일근무

    @Column(nullable = false)
    private Double holidayExtensionHours; // 휴일연장근무

    // 신청 월 (ex: "2026-04")
    @Column(nullable = false, length = 7)
    private String applyYearMonth;
}
