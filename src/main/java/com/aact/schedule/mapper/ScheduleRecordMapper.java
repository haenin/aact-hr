package com.aact.schedule.mapper;

import com.aact.schedule.dto.ScheduleRecordDto;
import com.aact.schedule.entity.ScheduleRecord;
import org.springframework.stereotype.Component;

import java.util.LinkedHashMap;
import java.util.Map;

@Component
public class ScheduleRecordMapper {

    public ScheduleRecord toEntity(ScheduleRecordDto.Request dto) {
        ScheduleRecord entity = ScheduleRecord.builder()
                .applyYearMonth(dto.getApplyYearMonth())
                .department(dto.getDepartment())
                .flightCode(dto.getFlightCode())
                .employeeName(dto.getEmployeeName())
                .seq(dto.getSeq())
                .offDays(dto.getOffDays())
                .usedOff(dto.getUsedOff())
                .usedAnnual(dto.getUsedAnnual())
                .remainAnnual(dto.getRemainAnnual())
                .build();

        if (dto.getDays() != null) {
            mapDaysToEntity(entity, dto.getDays());
        }
        return entity;
    }

    public ScheduleRecordDto.Response toResponse(ScheduleRecord entity) {
        return ScheduleRecordDto.Response.builder()
                .id(entity.getId())
                .applyYearMonth(entity.getApplyYearMonth())
                .department(entity.getDepartment())
                .flightCode(entity.getFlightCode())
                .employeeName(entity.getEmployeeName())
                .seq(entity.getSeq())
                .days(mapDaysFromEntity(entity))
                .offDays(entity.getOffDays())
                .usedOff(entity.getUsedOff())
                .usedAnnual(entity.getUsedAnnual())
                .remainAnnual(entity.getRemainAnnual())
                .build();
    }

    /** Map<Integer, String> → entity day01~day31 */
    private void mapDaysToEntity(ScheduleRecord e, Map<Integer, String> days) {
        e.setDay01(days.get(1));  e.setDay02(days.get(2));  e.setDay03(days.get(3));
        e.setDay04(days.get(4));  e.setDay05(days.get(5));  e.setDay06(days.get(6));
        e.setDay07(days.get(7));  e.setDay08(days.get(8));  e.setDay09(days.get(9));
        e.setDay10(days.get(10)); e.setDay11(days.get(11)); e.setDay12(days.get(12));
        e.setDay13(days.get(13)); e.setDay14(days.get(14)); e.setDay15(days.get(15));
        e.setDay16(days.get(16)); e.setDay17(days.get(17)); e.setDay18(days.get(18));
        e.setDay19(days.get(19)); e.setDay20(days.get(20)); e.setDay21(days.get(21));
        e.setDay22(days.get(22)); e.setDay23(days.get(23)); e.setDay24(days.get(24));
        e.setDay25(days.get(25)); e.setDay26(days.get(26)); e.setDay27(days.get(27));
        e.setDay28(days.get(28)); e.setDay29(days.get(29)); e.setDay30(days.get(30));
        e.setDay31(days.get(31));
    }

    /** entity day01~day31 → Map<Integer, String> */
    public Map<Integer, String> mapDaysFromEntity(ScheduleRecord e) {
        Map<Integer, String> days = new LinkedHashMap<>();
        days.put(1,  e.getDay01()); days.put(2,  e.getDay02()); days.put(3,  e.getDay03());
        days.put(4,  e.getDay04()); days.put(5,  e.getDay05()); days.put(6,  e.getDay06());
        days.put(7,  e.getDay07()); days.put(8,  e.getDay08()); days.put(9,  e.getDay09());
        days.put(10, e.getDay10()); days.put(11, e.getDay11()); days.put(12, e.getDay12());
        days.put(13, e.getDay13()); days.put(14, e.getDay14()); days.put(15, e.getDay15());
        days.put(16, e.getDay16()); days.put(17, e.getDay17()); days.put(18, e.getDay18());
        days.put(19, e.getDay19()); days.put(20, e.getDay20()); days.put(21, e.getDay21());
        days.put(22, e.getDay22()); days.put(23, e.getDay23()); days.put(24, e.getDay24());
        days.put(25, e.getDay25()); days.put(26, e.getDay26()); days.put(27, e.getDay27());
        days.put(28, e.getDay28()); days.put(29, e.getDay29()); days.put(30, e.getDay30());
        days.put(31, e.getDay31());
        // null인 날짜 제거 (해당 월에 없는 날)
        days.entrySet().removeIf(entry -> entry.getValue() == null);
        return days;
    }
}
