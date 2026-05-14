package com.aact.overtime.mapper;

import com.aact.overtime.dto.ApprovalLineDto;
import com.aact.overtime.entity.ApprovalLine;
import com.aact.overtime.type.ApprovalStatus;
import com.aact.overtime.type.ApproverRole;
import org.springframework.stereotype.Component;

import java.util.List;
import java.util.Map;
import java.util.function.Function;
import java.util.stream.Collectors;

@Component
public class ApprovalLineMapper {

    public ApprovalLine toEntity(String applyYearMonth, String department,
                                  ApprovalLineDto.ApproverItem item) {
        return ApprovalLine.builder()
                .applyYearMonth(applyYearMonth)
                .department(department)
                .role(item.getRole())
                .approverName(item.getApproverName())
                .status(ApprovalStatus.PENDING)
                .build();
    }

    public ApprovalLineDto.Response toResponse(ApprovalLine entity) {
        return ApprovalLineDto.Response.builder()
                .id(entity.getId())
                .applyYearMonth(entity.getApplyYearMonth())
                .department(entity.getDepartment())
                .role(entity.getRole())
                .approverName(entity.getApproverName())
                .status(entity.getStatus())
                .approvedAt(entity.getApprovedAt())
                .comment(entity.getComment())
                .build();
    }

    /** 결재라인 4명 → SummaryResponse */
    public ApprovalLineDto.SummaryResponse toSummary(String applyYearMonth,
                                                      String department,
                                                      List<ApprovalLine> lines) {
        Map<ApproverRole, ApprovalLine> map = lines.stream()
                .collect(Collectors.toMap(ApprovalLine::getRole, Function.identity()));

        // 전체 상태 계산: REJECTED 있으면 REJECTED, 전부 APPROVED면 APPROVED, 나머지는 PENDING
        ApprovalStatus overall;
        if (lines.stream().anyMatch(l -> l.getStatus() == ApprovalStatus.REJECTED)) {
            overall = ApprovalStatus.REJECTED;
        } else if (lines.stream().allMatch(l -> l.getStatus() == ApprovalStatus.APPROVED)) {
            overall = ApprovalStatus.APPROVED;
        } else {
            overall = ApprovalStatus.PENDING;
        }

        return ApprovalLineDto.SummaryResponse.builder()
                .applyYearMonth(applyYearMonth)
                .department(department)
                .manager(map.containsKey(ApproverRole.MANAGER) ? toResponse(map.get(ApproverRole.MANAGER)) : null)
                .chief(map.containsKey(ApproverRole.CHIEF) ? toResponse(map.get(ApproverRole.CHIEF)) : null)
                .director(map.containsKey(ApproverRole.DIRECTOR) ? toResponse(map.get(ApproverRole.DIRECTOR)) : null)
                .vp(map.containsKey(ApproverRole.VP) ? toResponse(map.get(ApproverRole.VP)) : null)
                .overallStatus(overall)
                .build();
    }
}
