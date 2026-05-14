package com.aact.schedule.service;

import com.aact.schedule.dto.ScheduleRecordDto;
import com.aact.schedule.entity.ScheduleRecord;
import com.aact.schedule.mapper.ScheduleRecordMapper;
import com.aact.schedule.repository.ScheduleRecordRepository;
import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.util.List;
import java.util.stream.Collectors;

@Service
@RequiredArgsConstructor
@Transactional(readOnly = true)
public class ScheduleRecordService {

    private final ScheduleRecordRepository scheduleRecordRepository;
    private final ScheduleRecordMapper scheduleRecordMapper;
    private final ScheduleExcelGenerator scheduleExcelGenerator;

    /** 단건 등록 */
    @Transactional
    public ScheduleRecordDto.Response create(ScheduleRecordDto.Request request) {
        ScheduleRecord entity = scheduleRecordMapper.toEntity(request);
        return scheduleRecordMapper.toResponse(scheduleRecordRepository.save(entity));
    }

    /** 일괄 등록 */
    @Transactional
    public List<ScheduleRecordDto.Response> createAll(List<ScheduleRecordDto.Request> requests) {
        List<ScheduleRecord> entities = requests.stream()
                .map(scheduleRecordMapper::toEntity)
                .collect(Collectors.toList());
        return scheduleRecordRepository.saveAll(entities).stream()
                .map(scheduleRecordMapper::toResponse)
                .collect(Collectors.toList());
    }

    /** 연월 + 부서 조회 */
    public List<ScheduleRecordDto.Response> getByYearMonthAndDepartment(
            String applyYearMonth, String department) {
        return scheduleRecordRepository
                .findByApplyYearMonthAndDepartmentOrderBySeqAsc(applyYearMonth, department)
                .stream()
                .map(scheduleRecordMapper::toResponse)
                .collect(Collectors.toList());
    }

    /** 단건 조회 */
    public ScheduleRecordDto.Response getById(Long id) {
        ScheduleRecord entity = scheduleRecordRepository.findById(id)
                .orElseThrow(() -> new IllegalArgumentException("해당 스케줄이 없습니다. id=" + id));
        return scheduleRecordMapper.toResponse(entity);
    }

    /** 수정 */
    @Transactional
    public ScheduleRecordDto.Response update(Long id, ScheduleRecordDto.Request request) {
        ScheduleRecord entity = scheduleRecordRepository.findById(id)
                .orElseThrow(() -> new IllegalArgumentException("해당 스케줄이 없습니다. id=" + id));

        entity.setDepartment(request.getDepartment());
        entity.setFlightCode(request.getFlightCode());
        entity.setEmployeeName(request.getEmployeeName());
        entity.setSeq(request.getSeq());
        entity.setOffDays(request.getOffDays());
        entity.setUsedOff(request.getUsedOff());
        entity.setUsedAnnual(request.getUsedAnnual());
        entity.setRemainAnnual(request.getRemainAnnual());

        if (request.getDays() != null) {
            ScheduleRecord temp = scheduleRecordMapper.toEntity(request);
            copyDays(temp, entity);
        }
        return scheduleRecordMapper.toResponse(entity);
    }

    /** 삭제 */
    @Transactional
    public void delete(Long id) {
        if (!scheduleRecordRepository.existsById(id)) {
            throw new IllegalArgumentException("해당 스케줄이 없습니다. id=" + id);
        }
        scheduleRecordRepository.deleteById(id);
    }

    /**
     * 엑셀 생성
     * 결재자, 편명, 팀명은 요청마다 동적으로 받음
     */
    public byte[] generateExcel(ScheduleRecordDto.ExcelRequest request) {
        try {
            return scheduleExcelGenerator.generate(
                    request.getApplyYearMonth(),
                    request.getDepartment(),
                    request.getFlightCode(),
                    request.getStatus() != null ? request.getStatus() : "확정",
                    request.getApprovers());
        } catch (Exception e) {
            throw new RuntimeException("스케줄 엑셀 생성 실패: " + e.getMessage(), e);
        }
    }

    private void copyDays(ScheduleRecord from, ScheduleRecord to) {
        to.setDay01(from.getDay01()); to.setDay02(from.getDay02()); to.setDay03(from.getDay03());
        to.setDay04(from.getDay04()); to.setDay05(from.getDay05()); to.setDay06(from.getDay06());
        to.setDay07(from.getDay07()); to.setDay08(from.getDay08()); to.setDay09(from.getDay09());
        to.setDay10(from.getDay10()); to.setDay11(from.getDay11()); to.setDay12(from.getDay12());
        to.setDay13(from.getDay13()); to.setDay14(from.getDay14()); to.setDay15(from.getDay15());
        to.setDay16(from.getDay16()); to.setDay17(from.getDay17()); to.setDay18(from.getDay18());
        to.setDay19(from.getDay19()); to.setDay20(from.getDay20()); to.setDay21(from.getDay21());
        to.setDay22(from.getDay22()); to.setDay23(from.getDay23()); to.setDay24(from.getDay24());
        to.setDay25(from.getDay25()); to.setDay26(from.getDay26()); to.setDay27(from.getDay27());
        to.setDay28(from.getDay28()); to.setDay29(from.getDay29()); to.setDay30(from.getDay30());
        to.setDay31(from.getDay31());
    }
}