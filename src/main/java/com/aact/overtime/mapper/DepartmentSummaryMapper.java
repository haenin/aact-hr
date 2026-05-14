package com.aact.overtime.mapper;

import com.aact.overtime.dto.DepartmentSummaryDto;
import com.aact.overtime.entity.DepartmentSummary;
import org.springframework.stereotype.Component;

@Component
public class DepartmentSummaryMapper {

    public DepartmentSummary toEntity(DepartmentSummaryDto.Request dto) {
        double currExt   = nvl(dto.getCurrExtensionHours());
        double currNight = nvl(dto.getCurrNightHours());
        double currHol   = nvl(dto.getCurrHolidayHours());
        double prevExt   = nvl(dto.getPrevExtensionHours());
        double prevNight = nvl(dto.getPrevNightHours());
        double prevHol   = nvl(dto.getPrevHolidayHours());

        return DepartmentSummary.builder()
                .department(dto.getDepartment())
                .applyYearMonth(dto.getApplyYearMonth())
                .prevExtensionHours(prevExt)
                .prevNightHours(prevNight)
                .prevHolidayHours(prevHol)
                .currExtensionHours(currExt)
                .currNightHours(currNight)
                .currHolidayHours(currHol)
                .diffExtensionHours(currExt - prevExt)
                .diffNightHours(currNight - prevNight)
                .diffHolidayHours(currHol - prevHol)
                .changeReason(dto.getChangeReason())
                .build();
    }

    public DepartmentSummaryDto.Response toResponse(DepartmentSummary entity) {
        return DepartmentSummaryDto.Response.builder()
                .id(entity.getId())
                .department(entity.getDepartment())
                .applyYearMonth(entity.getApplyYearMonth())
                .prevExtensionHours(entity.getPrevExtensionHours())
                .prevNightHours(entity.getPrevNightHours())
                .prevHolidayHours(entity.getPrevHolidayHours())
                .currExtensionHours(entity.getCurrExtensionHours())
                .currNightHours(entity.getCurrNightHours())
                .currHolidayHours(entity.getCurrHolidayHours())
                .diffExtensionHours(entity.getDiffExtensionHours())
                .diffNightHours(entity.getDiffNightHours())
                .diffHolidayHours(entity.getDiffHolidayHours())
                .changeReason(entity.getChangeReason())
                .build();
    }

    private double nvl(Double value) {
        return value != null ? value : 0.0;
    }
}
