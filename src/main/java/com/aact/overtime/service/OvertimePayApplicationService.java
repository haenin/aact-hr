package com.aact.overtime.service;

import com.aact.overtime.dto.OvertimePayApplicationDto;
import com.aact.overtime.entity.OvertimePayApplication;
import com.aact.overtime.mapper.OvertimePayApplicationMapper;
import com.aact.overtime.repository.OvertimePayApplicationRepository;
import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.util.List;
import java.util.stream.Collectors;

@Service
@RequiredArgsConstructor
@Transactional(readOnly = true)
public class OvertimePayApplicationService {

    private final OvertimePayApplicationRepository overtimeApplicationRepository;
    private final OvertimePayApplicationMapper overtimeApplicationMapper;
    private final OvertimePayApplicationExcelGenerator overtimeApplicationExcelGenerator;

    /** 단건 등록 */
    @Transactional
    public OvertimePayApplicationDto.Response create(OvertimePayApplicationDto.Request request) {
        OvertimePayApplication entity = overtimeApplicationMapper.toEntity(request);
        return overtimeApplicationMapper.toResponse(overtimeApplicationRepository.save(entity));
    }

    /** 일괄 등록 */
    @Transactional
    public List<OvertimePayApplicationDto.Response> createAll(List<OvertimePayApplicationDto.Request> requests) {
        List<OvertimePayApplication> entities = requests.stream()
                .map(overtimeApplicationMapper::toEntity)
                .collect(Collectors.toList());
        return overtimeApplicationRepository.saveAll(entities).stream()
                .map(overtimeApplicationMapper::toResponse)
                .collect(Collectors.toList());
    }

    /** 연월 + 부서 조회 */
    public List<OvertimePayApplicationDto.Response> getByYearMonthAndDepartment(
            String applyYearMonth, String department) {
        return overtimeApplicationRepository
                .findByApplyYearMonthAndDepartmentOrderByNoAsc(applyYearMonth, department)
                .stream()
                .map(overtimeApplicationMapper::toResponse)
                .collect(Collectors.toList());
    }

    /** 단건 조회 */
    public OvertimePayApplicationDto.Response getById(Long id) {
        OvertimePayApplication entity = overtimeApplicationRepository.findById(id)
                .orElseThrow(() -> new IllegalArgumentException("해당 신청서가 없습니다. id=" + id));
        return overtimeApplicationMapper.toResponse(entity);
    }

    /** 수정 */
    @Transactional
    public OvertimePayApplicationDto.Response update(Long id, OvertimePayApplicationDto.Request request) {
        OvertimePayApplication entity = overtimeApplicationRepository.findById(id)
                .orElseThrow(() -> new IllegalArgumentException("해당 신청서가 없습니다. id=" + id));

        entity.setNo(request.getNo());
        entity.setWorkDate(request.getWorkDate());
        entity.setEmployeeName(request.getEmployeeName());
        entity.setScheduledTime(request.getScheduledTime());
        entity.setActualTime(request.getActualTime());
        entity.setExtensionHours(request.getExtensionHours());
        entity.setNightHours(request.getNightHours());
        entity.setHolidayHours(request.getHolidayHours());
        entity.setWorkReason(request.getWorkReason());

        return overtimeApplicationMapper.toResponse(entity);
    }

    /** 삭제 */
    @Transactional
    public void delete(Long id) {
        if (!overtimeApplicationRepository.existsById(id)) {
            throw new IllegalArgumentException("해당 신청서가 없습니다. id=" + id);
        }
        overtimeApplicationRepository.deleteById(id);
    }

    /**
     * 엑셀 생성
     * - 결재란 직급, 서명자(신청자/팀장/부서장 등) 동적으로 받음
     */
    public byte[] generateExcel(OvertimePayApplicationDto.ExcelRequest request) {
        try {
            return overtimeApplicationExcelGenerator.generate(request);
        } catch (Exception e) {
            throw new RuntimeException("신청서 엑셀 생성 실패: " + e.getMessage(), e);
        }
    }
}
