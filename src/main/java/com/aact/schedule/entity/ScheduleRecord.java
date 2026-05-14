package com.aact.schedule.entity;

import jakarta.persistence.*;
import lombok.*;

/**
 * 월간 스케줄표 직원별 데이터
 * AACT {년}년 {월}월 {팀명} / {편명} 확정 SKD
 */
@Entity
@Table(name = "schedule_record",
       uniqueConstraints = @UniqueConstraint(columnNames = {"applyYearMonth", "department", "employeeName"}))
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class ScheduleRecord {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    @Column(nullable = false, length = 7)
    private String applyYearMonth;      // "2026-03"

    @Column(nullable = false, length = 50)
    private String department;          // 팀명 ex) "레약 & 교육품질팀"

    @Column(nullable = false, length = 50)
    private String flightCode;          // 편명 ex) "AA, 7L"

    @Column(nullable = false, length = 50)
    private String employeeName;        // 성명

    @Column(nullable = false)
    private Integer seq;                // 순번

    // 1일 ~ 31일 (없는 날짜는 null)
    // 값 종류: "X"(휴무), "9"(출근시간), "연"(연차), "반"(반차), "X/9"(혼합) 등
    private String day01; private String day02; private String day03;
    private String day04; private String day05; private String day06;
    private String day07; private String day08; private String day09;
    private String day10; private String day11; private String day12;
    private String day13; private String day14; private String day15;
    private String day16; private String day17; private String day18;
    private String day19; private String day20; private String day21;
    private String day22; private String day23; private String day24;
    private String day25; private String day26; private String day27;
    private String day28; private String day29; private String day30;
    private String day31;

    // 집계
    private Integer offDays;            // 휴무일수
    private Double usedOff;             // 사용휴무
    private Double usedAnnual;          // 사용연차
    private Double remainAnnual;        // 잔여연차
}
