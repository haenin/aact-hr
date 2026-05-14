package com.aact.overtime.repository;

import com.aact.overtime.entity.ApprovalLine;
import com.aact.overtime.type.ApprovalStatus;
import com.aact.overtime.type.ApproverRole;
import org.springframework.data.jpa.repository.JpaRepository;

import java.util.List;
import java.util.Optional;

public interface ApprovalLineRepository extends JpaRepository<ApprovalLine, Long> {

    /** 신청서 1건의 결재라인 전체 (담당→과장→상무→부사장) */
    List<ApprovalLine> findByApplyYearMonthAndDepartmentOrderByRoleAsc(
            String applyYearMonth, String department);

    /** 특정 결재자 단건 */
    Optional<ApprovalLine> findByApplyYearMonthAndDepartmentAndRole(
            String applyYearMonth, String department, ApproverRole role);

    /** 연월 전체 결재라인 (부서 무관) */
    List<ApprovalLine> findByApplyYearMonth(String applyYearMonth);

    /** 연월 전체 대기 중인 결재 */
    List<ApprovalLine> findByApplyYearMonthAndStatus(
            String applyYearMonth, ApprovalStatus status);

    /** 결재자 이름으로 본인 결재 목록 조회 */
    List<ApprovalLine> findByApproverNameAndStatus(
            String approverName, ApprovalStatus status);
}
