package com.aact.overtime.dto;

import com.aact.overtime.type.ApprovalStatus;
import com.aact.overtime.type.ApproverRole;
import lombok.*;
import java.time.LocalDateTime;
import java.util.List;

public class ApprovalLineDto {

    /** 결재라인 초기 생성 요청 (신청서 제출 시) */
    @Getter
    @Setter
    @NoArgsConstructor
    @AllArgsConstructor
    @Builder
    public static class CreateRequest {
        private String applyYearMonth;      // "2026-04"
        private String department;          // 부서명
        // 결재자 목록 (담당→과장→상무→부사장 순)
        private List<ApproverItem> approvers;
    }

    /** 결재자 한 명 정보 */
    @Getter
    @Setter
    @NoArgsConstructor
    @AllArgsConstructor
    @Builder
    public static class ApproverItem {
        private ApproverRole role;          // 담당 / 과장 / 상무 / 부사장
        private String approverName;        // 결재자 이름 (없으면 null)
    }

    /** 결재 처리 요청 (승인 or 반려) */
    @Getter
    @Setter
    @NoArgsConstructor
    @AllArgsConstructor
    @Builder
    public static class ApproveRequest {
        private ApprovalStatus status;      // APPROVED / REJECTED
        private String comment;             // 반려 사유 (반려 시 필수)
    }

    /** 결재라인 응답 */
    @Getter
    @Setter
    @NoArgsConstructor
    @AllArgsConstructor
    @Builder
    public static class Response {
        private Long id;
        private String applyYearMonth;
        private String department;
        private ApproverRole role;
        private String approverName;
        private ApprovalStatus status;      // PENDING / APPROVED / REJECTED
        private LocalDateTime approvedAt;
        private String comment;
    }

    /** 신청서 전체 결재 현황 (4명 한눈에) */
    @Getter
    @Setter
    @NoArgsConstructor
    @AllArgsConstructor
    @Builder
    public static class SummaryResponse {
        private String applyYearMonth;
        private String department;
        private Response manager;           // 담당
        private Response chief;             // 과장
        private Response director;          // 상무
        private Response vp;                // 부사장
        private ApprovalStatus overallStatus; // 전체 상태 (1명이라도 REJECTED면 REJECTED)
    }
}
