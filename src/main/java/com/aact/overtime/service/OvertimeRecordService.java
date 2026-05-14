package com.aact.overtime.service;

import com.aact.overtime.dto.OvertimeRecordDto;
import com.aact.overtime.entity.OvertimeRecord;
import com.aact.overtime.mapper.OvertimeRecordMapper;
import com.aact.overtime.repository.OvertimeRecordRepository;
import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.util.List;
import java.util.stream.Collectors;

@Service
@RequiredArgsConstructor
@Transactional(readOnly = true)
public class OvertimeRecordService {

    private final OvertimeRecordRepository overtimeRecordRepository;
    private final OvertimeRecordMapper overtimeRecordMapper;

    /** 단건 등록 */
    @Transactional
    public OvertimeRecordDto.Response create(OvertimeRecordDto.Request request) {
        OvertimeRecord entity = overtimeRecordMapper.toEntity(request);
        return overtimeRecordMapper.toResponse(overtimeRecordRepository.save(entity));
    }

    /** 일괄 등록 */
    @Transactional
    public List<OvertimeRecordDto.Response> createAll(List<OvertimeRecordDto.Request> requests) {
        List<OvertimeRecord> entities = requests.stream()
                .map(overtimeRecordMapper::toEntity)
                .collect(Collectors.toList());
        return overtimeRecordRepository.saveAll(entities).stream()
                .map(overtimeRecordMapper::toResponse)
                .collect(Collectors.toList());
    }

    /** 단건 조회 */
    public OvertimeRecordDto.Response getById(Long id) {
        OvertimeRecord entity = overtimeRecordRepository.findById(id)
                .orElseThrow(() -> new IllegalArgumentException("해당 시간외근무 내역이 없습니다. id=" + id));
        return overtimeRecordMapper.toResponse(entity);
    }

    /** 연월 전체 조회 */
    public List<OvertimeRecordDto.Response> getByYearMonth(String yearMonth) {
        return overtimeRecordRepository
                .findByApplyYearMonthOrderByDepartmentAscEmployeeNameAsc(yearMonth)
                .stream()
                .map(overtimeRecordMapper::toResponse)
                .collect(Collectors.toList());
    }

    /** 연월 + 부서 조회 */
    public List<OvertimeRecordDto.Response> getByYearMonthAndDepartment(String yearMonth, String department) {
        return overtimeRecordRepository
                .findByApplyYearMonthAndDepartmentOrderByWorkDateAsc(yearMonth, department)
                .stream()
                .map(overtimeRecordMapper::toResponse)
                .collect(Collectors.toList());
    }

    /** 연월 + 직원 조회 */
    public List<OvertimeRecordDto.Response> getByYearMonthAndEmployee(String yearMonth, String employeeName) {
        return overtimeRecordRepository
                .findByApplyYearMonthAndEmployeeNameOrderByWorkDateAsc(yearMonth, employeeName)
                .stream()
                .map(overtimeRecordMapper::toResponse)
                .collect(Collectors.toList());
    }

    /**
     * 직원별 소계 목록
     * ex) 부서 내 직원 각각의 연장/야간/휴일/휴일연장 합산
     */
    public List<OvertimeRecordDto.Subtotal> getSubtotalByDepartment(String yearMonth, String department) {
        return overtimeRecordRepository.aggregateByEmployee(yearMonth, department)
                .stream()
                .map(row -> OvertimeRecordDto.Subtotal.builder()
                        .employeeName((String) row[0])
                        .extensionHours(toDouble(row[1]))
                        .nightHours(toDouble(row[2]))
                        .holidayHours(toDouble(row[3]))
                        .holidayExtensionHours(toDouble(row[4]))
                        .build())
                .collect(Collectors.toList());
    }

    /**
     * 부서 전체 합계
     * ex) 항공기운송지원(운항팀) 합계 행
     */
    public OvertimeRecordDto.Total getTotalByDepartment(String yearMonth, String department) {
        return overtimeRecordRepository.aggregateByDepartment(yearMonth)
                .stream()
                .filter(row -> department.equals(row[0]))
                .findFirst()
                .map(row -> OvertimeRecordDto.Total.builder()
                        .department((String) row[0])
                        .extensionHours(toDouble(row[1]))
                        .nightHours(toDouble(row[2]))
                        .holidayHours(toDouble(row[3]))
                        .holidayExtensionHours(toDouble(row[4]))
                        .build())
                .orElse(OvertimeRecordDto.Total.builder()
                        .department(department)
                        .extensionHours(0.0)
                        .nightHours(0.0)
                        .holidayHours(0.0)
                        .holidayExtensionHours(0.0)
                        .build());
    }

    /** 수정 */
    @Transactional
    public OvertimeRecordDto.Response update(Long id, OvertimeRecordDto.Request request) {
        OvertimeRecord entity = overtimeRecordRepository.findById(id)
                .orElseThrow(() -> new IllegalArgumentException("해당 시간외근무 내역이 없습니다. id=" + id));

        entity.setDepartment(request.getDepartment());
        entity.setEmployeeName(request.getEmployeeName());
        entity.setWorkDate(request.getWorkDate());
        entity.setScheduledStart(request.getScheduledStart());
        entity.setScheduledEnd(request.getScheduledEnd());
        entity.setActualStart(request.getActualStart());
        entity.setActualEnd(request.getActualEnd());
        entity.setExtensionHours(request.getExtensionHours());
        entity.setNightHours(request.getNightHours());
        entity.setHolidayHours(request.getHolidayHours());
        entity.setHolidayExtensionHours(request.getHolidayExtensionHours());

        return overtimeRecordMapper.toResponse(entity);
    }

    /** 삭제 */
    @Transactional
    public void delete(Long id) {
        if (!overtimeRecordRepository.existsById(id)) {
            throw new IllegalArgumentException("해당 시간외근무 내역이 없습니다. id=" + id);
        }
        overtimeRecordRepository.deleteById(id);
    }

    private double toDouble(Object v) {
        return v != null ? ((Number) v).doubleValue() : 0.0;
    }
}