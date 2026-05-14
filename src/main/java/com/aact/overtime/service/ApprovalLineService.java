package com.aact.overtime.service;

import com.aact.overtime.dto.ApprovalLineDto;
import com.aact.overtime.entity.ApprovalLine;
import com.aact.overtime.mapper.ApprovalLineMapper;
import com.aact.overtime.repository.ApprovalLineRepository;
import com.aact.overtime.type.ApprovalStatus;
import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.time.LocalDateTime;
import java.util.List;
import java.util.stream.Collectors;

@Service
@RequiredArgsConstructor
@Transactional(readOnly = true)
public class ApprovalLineService {

    private final ApprovalLineRepository approvalLineRepository;
    private final ApprovalLineMapper approvalLineMapper;

    /**
     * 결재라인 생성 (신청서 제출 시 담당/과장/상무/부사장 4명 한번에 생성)
     * 결재자 이름은 없어도 됨 (나중에 업데이트 가능)
     */
    @Transactional
    public ApprovalLineDto.SummaryResponse create(ApprovalLineDto.CreateRequest request) {
        List<ApprovalLine> lines = request.getApprovers().stream()
                .map(item -> approvalLineMapper.toEntity(
                        request.getApplyYearMonth(),
                        request.getDepartment(),
                        item))
                .collect(Collectors.toList());

        List<ApprovalLine> saved = approvalLineRepository.saveAll(lines);
        return approvalLineMapper.toSummary(
                request.getApplyYearMonth(),
                request.getDepartment(),
                saved);
    }

    /**
     * 결재 현황 조회 (신청서 1건의 담당/과장/상무/부사장 전체)
     */
    public ApprovalLineDto.SummaryResponse getSummary(String applyYearMonth, String department) {
        List<ApprovalLine> lines = approvalLineRepository
                .findByApplyYearMonthAndDepartmentOrderByRoleAsc(applyYearMonth, department);

        if (lines.isEmpty()) {
            throw new IllegalArgumentException("결재라인이 없습니다. " + applyYearMonth + " / " + department);
        }
        return approvalLineMapper.toSummary(applyYearMonth, department, lines);
    }

    /**
     * 결재 처리 (승인 or 반려)
     * id = ApprovalLine 단건 id
     */
    @Transactional
    public ApprovalLineDto.Response approve(Long id, ApprovalLineDto.ApproveRequest request) {
        ApprovalLine line = approvalLineRepository.findById(id)
                .orElseThrow(() -> new IllegalArgumentException("결재 항목이 없습니다. id=" + id));

        if (line.getStatus() != ApprovalStatus.PENDING) {
            throw new IllegalStateException("이미 처리된 결재입니다. status=" + line.getStatus());
        }
        if (request.getStatus() == ApprovalStatus.REJECTED && 
            (request.getComment() == null || request.getComment().isBlank())) {
            throw new IllegalArgumentException("반려 시 사유를 입력해야 합니다.");
        }

        line.setStatus(request.getStatus());
        line.setApprovedAt(LocalDateTime.now());
        line.setComment(request.getComment());

        return approvalLineMapper.toResponse(line);
    }

    /**
     * 결재자 이름 업데이트 (결재자가 바뀔 때)
     */
    @Transactional
    public ApprovalLineDto.Response updateApprover(Long id, String approverName) {
        ApprovalLine line = approvalLineRepository.findById(id)
                .orElseThrow(() -> new IllegalArgumentException("결재 항목이 없습니다. id=" + id));

        line.setApproverName(approverName);
        return approvalLineMapper.toResponse(line);
    }

    /**
     * 내 결재 대기 목록 (결재자 이름으로 조회)
     */
    public List<ApprovalLineDto.Response> getPendingList(String approverName) {
        return approvalLineRepository
                .findByApproverNameAndStatus(approverName, ApprovalStatus.PENDING)
                .stream()
                .map(approvalLineMapper::toResponse)
                .collect(Collectors.toList());
    }
}
