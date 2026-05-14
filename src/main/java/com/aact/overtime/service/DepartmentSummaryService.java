package com.aact.overtime.service;

import com.aact.overtime.dto.DepartmentSummaryDto;
import com.aact.overtime.entity.DepartmentSummary;
import com.aact.overtime.mapper.DepartmentSummaryMapper;
import com.aact.overtime.repository.DepartmentSummaryRepository;
import com.aact.overtime.repository.OvertimeRecordRepository;
import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.util.List;
import java.util.stream.Collectors;

@Service
@RequiredArgsConstructor
@Transactional(readOnly = true)
public class DepartmentSummaryService {

    private final DepartmentSummaryRepository departmentSummaryRepository;
    private final OvertimeRecordRepository overtimeRecordRepository;
    private final DepartmentSummaryMapper departmentSummaryMapper;

    /** 집계표 수동 저장/수정 (표지 입력) */
    @Transactional
    public DepartmentSummaryDto.Response saveOrUpdate(DepartmentSummaryDto.Request request) {
        DepartmentSummary entity = departmentSummaryRepository
                .findByDepartmentAndApplyYearMonth(request.getDepartment(), request.getApplyYearMonth())
                .orElse(null);

        if (entity == null) {
            entity = departmentSummaryMapper.toEntity(request);
        } else {
            // 기존 값 업데이트
            entity.setPrevExtensionHours(nvl(request.getPrevExtensionHours()));
            entity.setPrevNightHours(nvl(request.getPrevNightHours()));
            entity.setPrevHolidayHours(nvl(request.getPrevHolidayHours()));
            entity.setCurrExtensionHours(nvl(request.getCurrExtensionHours()));
            entity.setCurrNightHours(nvl(request.getCurrNightHours()));
            entity.setCurrHolidayHours(nvl(request.getCurrHolidayHours()));
            entity.setDiffExtensionHours(nvl(request.getCurrExtensionHours()) - nvl(request.getPrevExtensionHours()));
            entity.setDiffNightHours(nvl(request.getCurrNightHours()) - nvl(request.getPrevNightHours()));
            entity.setDiffHolidayHours(nvl(request.getCurrHolidayHours()) - nvl(request.getPrevHolidayHours()));
            entity.setChangeReason(request.getChangeReason());
        }

        return departmentSummaryMapper.toResponse(departmentSummaryRepository.save(entity));
    }

    /** 연월 전체 부서 집계 조회 */
    public List<DepartmentSummaryDto.Response> getByYearMonth(String yearMonth) {
        return departmentSummaryRepository
                .findByApplyYearMonthOrderByDepartmentAsc(yearMonth)
                .stream()
                .map(departmentSummaryMapper::toResponse)
                .collect(Collectors.toList());
    }

    /**
     * 개인별 내역 기반으로 당월 부서 집계를 자동 계산
     * → OvertimeRecord의 합산값으로 DepartmentSummary의 curr* 필드를 갱신
     */
    @Transactional
    public List<DepartmentSummaryDto.AggregateResult> recalculate(String yearMonth) {
        List<Object[]> rows = overtimeRecordRepository.aggregateByDepartment(yearMonth);

        return rows.stream().map(row -> {
            String dept      = (String) row[0];
            double ext       = toDouble(row[1]);
            double night     = toDouble(row[2]);
            double holiday   = toDouble(row[3]);
            double holExt    = toDouble(row[4]);

            // DepartmentSummary curr* 자동 업데이트
            departmentSummaryRepository
                    .findByDepartmentAndApplyYearMonth(dept, yearMonth)
                    .ifPresent(summary -> {
                        double diff;
                        summary.setCurrExtensionHours(ext);
                        summary.setCurrNightHours(night);
                        summary.setCurrHolidayHours(holiday);
                        summary.setDiffExtensionHours(ext - nvl(summary.getPrevExtensionHours()));
                        summary.setDiffNightHours(night - nvl(summary.getPrevNightHours()));
                        summary.setDiffHolidayHours(holiday - nvl(summary.getPrevHolidayHours()));
                        departmentSummaryRepository.save(summary);
                    });

            return DepartmentSummaryDto.AggregateResult.builder()
                    .department(dept)
                    .totalExtension(ext)
                    .totalNight(night)
                    .totalHoliday(holiday)
                    .totalHolidayExtension(holExt)
                    .build();
        }).collect(Collectors.toList());
    }

    private double nvl(Double v) { return v != null ? v : 0.0; }
    private double toDouble(Object v) { return v != null ? ((Number) v).doubleValue() : 0.0; }
}
