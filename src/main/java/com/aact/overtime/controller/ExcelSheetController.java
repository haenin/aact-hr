package com.aact.overtime.controller;

import com.aact.overtime.dto.SheetDto;
import lombok.RequiredArgsConstructor;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;

/**
 * 엑셀 → 시트 배열 변환 API
 *
 * POST /api/overtime/excel/parse
 * multipart/form-data: file=시간외근무수당신청서.xlsx
 *
 * 응답:
 * [
 *   {
 *     "sheetName": "조업부표지",
 *     "isCover": true,
 *     "type": "조업부",
 *     "data": [
 *       {
 *         "departmentName": "조업부",
 *         "prevExtension": 0.0,
 *         "prevNight": 0.0,
 *         "prevHoliday": 0.0,
 *         "currExtension": 0.0,
 *         "currNight": 0.0,
 *         "currHoliday": 0.0,
 *         "diffExtension": 0.0,
 *         "diffNight": 0.0,
 *         "diffHoliday": 0.0,
 *         "changeReason": null
 *       },
 *       ...
 *     ]
 *   },
 *   {
 *     "sheetName": "항공기운송지원(운항팀)",
 *     "isCover": false,
 *     "type": "항공기운송지원",
 *     "data": [
 *       {
 *         "department": "운항팀",
 *         "employeeName": "김보라",
 *         "workDate": "2026-05-01",
 *         "scheduledStart": "09:00",
 *         "scheduledEnd": "18:00",
 *         "actualStart": "09:00",
 *         "actualEnd": "20:30",
 *         "extensionHours": 2.5,
 *         "nightHours": 0.0,
 *         "holidayHours": 0.0,
 *         "holidayExtensionHours": 0.0,
 *         "isSubtotal": false
 *       },
 *       ...
 *     ]
 *   }
 * ]
 */
@RestController
@RequestMapping("/api/overtime/excel")
@RequiredArgsConstructor
public class ExcelSheetController {

    private final ExcelSheetService excelSheetService;

    @PostMapping("/parse")
    public ResponseEntity<List<SheetDto.SheetWrapper<?>>> parse(
            @RequestParam("file") MultipartFile file) {

        if (file.isEmpty()) {
            return ResponseEntity.badRequest().build();
        }
        return ResponseEntity.ok(excelSheetService.parseExcel(file));
    }
}
