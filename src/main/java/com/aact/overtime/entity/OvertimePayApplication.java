package com.aact.overtime.entity;

import jakarta.persistence.*;
import lombok.*;

/**
 * 시간외근무수당 신청서 - 개인별 행 데이터
 * 시간 외 근무수당 신청서(관리부) 양식
 */
@Entity
@Table(name = "overtime_application")
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class OvertimePayApplication {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    @Column(nullable = false, length = 7)
    private String applyYearMonth;      // "2026-03"

    @Column(nullable = false, length = 50)
    private String department;          // 부서명 ex) "관리부"

    @Column(nullable = false)
    private Integer no;                 // NO (순번)

    @Column(length = 20)
    private String workDate;            // 일자 ex) "03월05일"

    @Column(length = 50)
    private String employeeName;        // 성명

    @Column(length = 30)
    private String scheduledTime;       // 예정근무시간 ex) "10:00~19:00"

    @Column(length = 30)
    private String actualTime;          // 실 근무시간 ex) "10:00~22:30"

    private Double extensionHours;      // 연장근무시간
    private Double nightHours;          // 야간근무시간
    private Double holidayHours;        // 휴일근무시간

    @Column(length = 200)
    private String workReason;          // 근무사유
}
