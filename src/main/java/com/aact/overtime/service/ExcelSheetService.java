package com.aact.overtime.service;

import com.aact.overtime.dto.SheetDto;
import com.aact.overtime.dto.SheetType;
import com.aact.overtime.service.ExcelSheetParser;
import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;
import java.util.stream.Collectors;

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

    /**
     * 표지 시트만 반환 (isCover = true)
     * → 부서별 집계표 (전월/당월/증감)
     */
    public List<SheetDto.SheetWrapper<?>> parseCoverSheets(MultipartFile file) {
        return parseExcel(file).stream()
                .filter(SheetDto.SheetWrapper::isCover)
                .collect(Collectors.toList());
    }

    /**
     * 세부내역 시트만 반환 (isCover = false)
     * → 개인별 일자별 근무내역
     */
    public List<SheetDto.SheetWrapper<?>> parseDetailSheets(MultipartFile file) {
        return parseExcel(file).stream()
                .filter(sheet -> !sheet.isCover())
                .collect(Collectors.toList());
    }

    /**
     * 특정 부서 타입 시트만 반환
     * ex) SheetType.조업부 → 조업부표지 + 조업부 세부내역
     */
    public List<SheetDto.SheetWrapper<?>> parseByType(MultipartFile file, SheetType type) {
        return parseExcel(file).stream()
                .filter(sheet -> sheet.getType() == type)
                .collect(Collectors.toList());
    }
}
