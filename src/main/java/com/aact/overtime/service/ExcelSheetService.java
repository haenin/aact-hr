package com.aact.overtime.service;

import com.aact.overtime.dto.SheetDto;
import com.aact.overtime.service.ExcelSheetParser;
import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;

@Service
@RequiredArgsConstructor
public class ExcelSheetService {

    private final ExcelSheetParser excelSheetParser;

    /**
     * 업로드된 엑셀 파일을 시트 배열로 변환
     *
     * 반환 예시:
     * [
     *   { sheetName: "조업부표지",           isCover: true,  type: 조업부,         data: [...CoverRow] },
     *   { sheetName: "항공기운송지원표지",     isCover: true,  type: 항공기운송지원,  data: [...CoverRow] },
     *   { sheetName: "관리부표지",            isCover: true,  type: 관리부,         data: [...CoverRow] },
     *   { sheetName: "항공기운송지원(운항팀)", isCover: false, type: 항공기운송지원,  data: [...DetailRow] }
     * ]
     */
    public List<SheetDto.SheetWrapper<?>> parseExcel(MultipartFile file) {
        try {
            return excelSheetParser.parse(file.getInputStream());
        } catch (Exception e) {
            throw new RuntimeException("엑셀 파싱 실패: " + e.getMessage(), e);
        }
    }
}
