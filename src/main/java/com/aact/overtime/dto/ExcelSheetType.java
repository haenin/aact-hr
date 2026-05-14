package com.aact.overtime.dto;

/**
 * 시트 종류
 * 엑셀 시트명 패턴으로 자동 판별
 */
public enum ExcelSheetType {
    조업부,
    항공기운송지원,
    관리부;

    /**
     * 시트명에서 타입 추론
     * ex) "조업부표지" -> 조업부, "항공기운송지원(운항팀)" -> 항공기운송지원
     */
    public static ExcelSheetType from(String sheetName) {
        for (ExcelSheetType type : values()) {
            if (sheetName.contains(type.name())) {
                return type;
            }
        }
        throw new IllegalArgumentException("알 수 없는 시트 타입: " + sheetName);
    }
}
