//package com.aact.overtime.controller;
//
//import com.aact.overtime.dto.DepartmentSummaryDto;
//import com.aact.overtime.service.DepartmentSummaryService;
//import lombok.RequiredArgsConstructor;
//import org.springframework.http.ResponseEntity;
//import org.springframework.web.bind.annotation.*;
//
//import java.util.List;
//
///**
// * 부서별 집계표 API
// * Base URL: /api/overtime/summaries
// */
//@RestController
//@RequestMapping("/api/overtime/summaries")
//@RequiredArgsConstructor
//public class DepartmentSummaryController {
//
//    private final DepartmentSummaryService departmentSummaryService;
//
//    /** POST /api/overtime/summaries — 집계표 저장/수정 */
//    @PostMapping
//    public ResponseEntity<DepartmentSummaryDto.Response> saveOrUpdate(
//            @RequestBody DepartmentSummaryDto.Request request) {
//        return ResponseEntity.ok(departmentSummaryService.saveOrUpdate(request));
//    }
//
//    /** GET /api/overtime/summaries?yearMonth=2026-04 — 연월 전체 조회 */
//    @GetMapping
//    public ResponseEntity<List<DepartmentSummaryDto.Response>> getByYearMonth(
//            @RequestParam String yearMonth) {
//        return ResponseEntity.ok(departmentSummaryService.getByYearMonth(yearMonth));
//    }
//
//    /**
//     * POST /api/overtime/summaries/recalculate?yearMonth=2026-04
//     * 개인별 내역 합산 → 당월 집계 자동 갱신
//     */
//    @PostMapping("/recalculate")
//    public ResponseEntity<List<DepartmentSummaryDto.AggregateResult>> recalculate(
//            @RequestParam String yearMonth) {
//        return ResponseEntity.ok(departmentSummaryService.recalculate(yearMonth));
//    }
//}
