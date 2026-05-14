package com.aact.overtime.service;

import com.aact.overtime.dto.ExcelSheetDto;
import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;

@Service
@RequiredArgsConstructor
public class ExcelSheetService {

    private final ExcelSheetParser excelSheetParser;
    private final ExcelSheetGenerator excelSheetGenerator;

    /**
     * 엑셀 파싱 → 시트 배열 반환
     */
    public List<ExcelSheetDto.SheetWrapper<?>> parseExcel(MultipartFile file) {
        try {
            return excelSheetParser.parse(file.getInputStream());
        } catch (Exception e) {
            throw new RuntimeException("엑셀 파싱 실패: " + e.getMessage(), e);
        }
    }

    /**
     * 엑셀 생성
     * @param applyYearMonth "2026-04"
     * @param approvers      결재란 직급 목록 ex) ["담당", "과장", "상무", "부사장"]
     */
    public byte[] generateExcel(String applyYearMonth, List<String> approvers) {
        try {
            return excelSheetGenerator.generate(applyYearMonth, approvers);
        } catch (Exception e) {
            throw new RuntimeException("엑셀 생성 실패: " + e.getMessage(), e);
        }
    }
}
