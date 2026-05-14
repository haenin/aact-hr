package com.aact.overtime.dto;

import lombok.*;

import java.util.List;

/**
 * 엑셀 전체를 시트 배열로 표현하는 최상위 DTO
 *
 * 구조 예시:
 * [
 *   {
 *     "sheetName": "조업부표지",
 *     "isCover": true,
 *     "type": "조업부",
 *     "data": [ { CoverRowDto } ]
 *   },
 *   {
 *     "sheetName": "항공기운송지원(운항팀)",
 *     "isCover": false,
 *     "type": "항공기운송지원",
 *     "data": [ { DetailRowDto } ]
 *   }
 * ]
 */
public class ExcelSheetDto {

    // ────────────────────────────────────────────
    // 표지 시트용 (isCover = true)
    // ────────────────────────────────────────────

    @Getter
    @Setter
    @NoArgsConstructor
    @AllArgsConstructor
    @Builder
    public static class CoverSheet {
        private String sheetName;   // 시트명
        private boolean isCover;    // 표지 여부 (항상 true)
        private ExcelSheetType type;     // 부서 타입 (조업부 / 항공기운송지원 / 관리부)
        private List<CoverRow> data;
    }

    @Getter
    @Setter
    @NoArgsConstructor
    @AllArgsConstructor
    @Builder
    public static class CoverRow {
        private String departmentName;      // 부서명
        private Double prevExtension;       // 전월 연장
        private Double prevNight;           // 전월 야간
        private Double prevHoliday;         // 전월 휴일
        private Double currExtension;       // 당월 연장
        private Double currNight;           // 당월 야간
        private Double currHoliday;         // 당월 휴일
        private Double diffExtension;       // 증감 연장
        private Double diffNight;           // 증감 야간
        private Double diffHoliday;         // 증감 휴일
        private String changeReason;        // 증감사유 / OT사유
    }

    // ────────────────────────────────────────────
    // 개인별 세부내역 시트용 (isCover = false)
    // ────────────────────────────────────────────

    @Getter
    @Setter
    @NoArgsConstructor
    @AllArgsConstructor
    @Builder
    public static class DetailSheet {
        private String sheetName;   // 시트명
        private boolean isCover;    // 표지 여부 (항상 false)
        private ExcelSheetType type;     // 부서 타입
        private List<DetailRow> data;
    }

    @Getter
    @Setter
    @NoArgsConstructor
    @AllArgsConstructor
    @Builder
    public static class DetailRow {
        private String department;          // 부서
        private String employeeName;        // 성명
        private String workDate;            // 일자 (yyyy-MM-dd)
        private String scheduledStart;      // 예정 출근
        private String scheduledEnd;        // 예정 퇴근
        private String actualStart;         // 실근무 출근 (지문)
        private String actualEnd;           // 실근무 퇴근 (지문)
        private Double extensionHours;      // 연장
        private Double nightHours;          // 야간
        private Double holidayHours;        // 휴일
        private Double holidayExtensionHours; // 휴일연장
        private boolean isSubtotal;         // 소계 행 여부
    }

    // ────────────────────────────────────────────
    // 공통 래퍼 (표지/세부 구분 없이 배열로 쓸 때)
    // ────────────────────────────────────────────

    @Getter
    @Setter
    @NoArgsConstructor
    @AllArgsConstructor
    @Builder
    public static class SheetWrapper<T> {
        private String sheetName;
        private boolean isCover;
        private ExcelSheetType type;
        private List<T> data;
    }
}
