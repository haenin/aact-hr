package com.aact.overtime.service;

import com.aact.overtime.dto.ApplicationDto;
import com.aact.overtime.entity.OvertimePayApplication;
import com.aact.overtime.mapper.ApplicationMapper;
import com.aact.overtime.repository.ApplicationRepository;
import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;
import java.util.stream.Collectors;

@Service
@RequiredArgsConstructor
@Transactional(readOnly = true)
public class ApplicationService {

    private final ApplicationRepository overtimeApplicationRepository;
    private final ApplicationMapper overtimeApplicationMapper;
    private final ApplicationExcelGenerator applicationExcelGenerator;
    private final ApplicationExcelParser applicationExcelParser;

    /** 단건 등록 */
    @Transactional
    public ApplicationDto.Response create(ApplicationDto.Request request) {
        OvertimePayApplication entity = overtimeApplicationMapper.toEntity(request);
        return overtimeApplicationMapper.toResponse(overtimeApplicationRepository.save(entity));
    }

    /** 일괄 등록 */
    @Transactional
    public List<ApplicationDto.Response> createAll(List<ApplicationDto.Request> requests) {
        List<OvertimePayApplication> entities = requests.stream()
                .map(overtimeApplicationMapper::toEntity)
                .collect(Collectors.toList());
        return overtimeApplicationRepository.saveAll(entities).stream()
                .map(overtimeApplicationMapper::toResponse)
                .collect(Collectors.toList());
    }

    /** 연월 + 부서 조회 */
    public List<ApplicationDto.Response> getByYearMonthAndDepartment(
            String applyYearMonth, String department) {
        return overtimeApplicationRepository
                .findByApplyYearMonthAndDepartmentOrderByNoAsc(applyYearMonth, department)
                .stream()
                .map(overtimeApplicationMapper::toResponse)
                .collect(Collectors.toList());
    }

    /** 단건 조회 */
    public ApplicationDto.Response getById(Long id) {
        OvertimePayApplication entity = overtimeApplicationRepository.findById(id)
                .orElseThrow(() -> new IllegalArgumentException("해당 신청서가 없습니다. id=" + id));
        return overtimeApplicationMapper.toResponse(entity);
    }

    /** 수정 */
    @Transactional
    public ApplicationDto.Response update(Long id, ApplicationDto.Request request) {
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
     * 엑셀 파싱 → DB 저장
     * 엑셀 업로드하면 파싱해서 바로 저장
     */
    @Transactional
    public List<ApplicationDto.Response> parseAndSave(
            MultipartFile file, String applyYearMonth, String department) {
        try {
            List<ApplicationDto.Request> requests =
                    applicationExcelParser.parse(file.getInputStream(), applyYearMonth, department);
            return createAll(requests);
        } catch (Exception e) {
            throw new RuntimeException("신청서 파싱 실패: " + e.getMessage(), e);
        }
    }

    /**
     * 엑셀 파싱만 (저장 없이 미리보기용)
     */
    public List<ApplicationDto.Request> parseOnly(
            MultipartFile file, String applyYearMonth, String department) {
        try {
            return applicationExcelParser.parse(file.getInputStream(), applyYearMonth, department);
        } catch (Exception e) {
            throw new RuntimeException("신청서 파싱 실패: " + e.getMessage(), e);
        }
    }

    /**
     * 엑셀 생성
     * 결재란, 서명자 동적으로 받음
     */
    public byte[] generateExcel(ApplicationDto.ExcelRequest request) {
        try {
            return applicationExcelGenerator.generate(request);
        } catch (Exception e) {
            throw new RuntimeException("신청서 엑셀 생성 실패: " + e.getMessage(), e);
        }
    }
}