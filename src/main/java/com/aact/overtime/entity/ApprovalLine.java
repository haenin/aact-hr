package com.aact.overtime.entity;

import com.aact.overtime.type.ApprovalStatus;
import com.aact.overtime.type.ApproverRole;
import jakarta.persistence.*;
import lombok.*;
import java.time.LocalDateTime;

/**
 * 결재라인
 * 시간외근무 신청서 1건당 결재자 N명 (담당/과장/상무/부사장)
 */
@Entity
@Table(name = "approval_line")
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class ApprovalLine {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    // 어떤 신청서에 대한 결재인지 (신청 연월 + 부서로 식별)
    @Column(nullable = false, length = 7)
    private String applyYearMonth;      // ex) "2026-04"

    @Column(nullable = false, length = 50)
    private String department;          // 부서명

    // 결재자 정보
    @Enumerated(EnumType.STRING)
    @Column(nullable = false)
    private ApproverRole role;          // 담당 / 과장 / 상무 / 부사장

    @Column(length = 50)
    private String approverName;        // 결재자 이름 (변경될 수 있어서 nullable)

    // 결재 상태
    @Enumerated(EnumType.STRING)
    @Column(nullable = false)
    @Builder.Default
    private ApprovalStatus status = ApprovalStatus.PENDING;  // 기본값: 대기

    private LocalDateTime approvedAt;   // 결재 일시 (승인/반려 시 기록)

    @Column(length = 500)
    private String comment;             // 반려 사유 등
}
