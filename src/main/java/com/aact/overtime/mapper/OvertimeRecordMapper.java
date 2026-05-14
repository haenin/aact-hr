package com.aact.overtime.mapper;

import com.aact.overtime.dto.OvertimeRecordDto;
import com.aact.overtime.entity.OvertimeRecord;
import org.springframework.stereotype.Component;

@Component
public class OvertimeRecordMapper {

    public OvertimeRecord toEntity(OvertimeRecordDto.Request dto) {
        return OvertimeRecord.builder()
                .department(dto.getDepartment())
                .employeeName(dto.getEmployeeName())
                .workDate(dto.getWorkDate())
                .scheduledStart(dto.getScheduledStart())
                .scheduledEnd(dto.getScheduledEnd())
                .actualStart(dto.getActualStart())
                .actualEnd(dto.getActualEnd())
                .extensionHours(dto.getExtensionHours() != null ? dto.getExtensionHours() : 0.0)
                .nightHours(dto.getNightHours() != null ? dto.getNightHours() : 0.0)
                .holidayHours(dto.getHolidayHours() != null ? dto.getHolidayHours() : 0.0)
                .holidayExtensionHours(dto.getHolidayExtensionHours() != null ? dto.getHolidayExtensionHours() : 0.0)
                .applyYearMonth(dto.getApplyYearMonth())
                .build();
    }

    public OvertimeRecordDto.Response toResponse(OvertimeRecord entity) {
        return OvertimeRecordDto.Response.builder()
                .id(entity.getId())
                .department(entity.getDepartment())
                .employeeName(entity.getEmployeeName())
                .workDate(entity.getWorkDate())
                .scheduledStart(entity.getScheduledStart())
                .scheduledEnd(entity.getScheduledEnd())
                .actualStart(entity.getActualStart())
                .actualEnd(entity.getActualEnd())
                .extensionHours(entity.getExtensionHours())
                .nightHours(entity.getNightHours())
                .holidayHours(entity.getHolidayHours())
                .holidayExtensionHours(entity.getHolidayExtensionHours())
                .applyYearMonth(entity.getApplyYearMonth())
                .build();
    }
}
