package com.aact.overtime.mapper;

import com.aact.overtime.dto.ApplicationDto;
import com.aact.overtime.entity.OvertimePayApplication;
import org.springframework.stereotype.Component;

@Component
public class ApplicationMapper {

    public OvertimePayApplication toEntity(ApplicationDto.Request dto) {
        return OvertimePayApplication.builder()
                .applyYearMonth(dto.getApplyYearMonth())
                .department(dto.getDepartment())
                .no(dto.getNo())
                .workDate(dto.getWorkDate())
                .employeeName(dto.getEmployeeName())
                .scheduledTime(dto.getScheduledTime())
                .actualTime(dto.getActualTime())
                .extensionHours(dto.getExtensionHours() != null ? dto.getExtensionHours() : 0.0)
                .nightHours(dto.getNightHours() != null ? dto.getNightHours() : 0.0)
                .holidayHours(dto.getHolidayHours() != null ? dto.getHolidayHours() : 0.0)
                .workReason(dto.getWorkReason())
                .build();
    }

    public ApplicationDto.Response toResponse(OvertimePayApplication entity) {
        return ApplicationDto.Response.builder()
                .id(entity.getId())
                .applyYearMonth(entity.getApplyYearMonth())
                .department(entity.getDepartment())
                .no(entity.getNo())
                .workDate(entity.getWorkDate())
                .employeeName(entity.getEmployeeName())
                .scheduledTime(entity.getScheduledTime())
                .actualTime(entity.getActualTime())
                .extensionHours(entity.getExtensionHours())
                .nightHours(entity.getNightHours())
                .holidayHours(entity.getHolidayHours())
                .workReason(entity.getWorkReason())
                .build();
    }
}
