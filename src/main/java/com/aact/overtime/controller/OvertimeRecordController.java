package com.aact.overtime.controller;

import com.aact.overtime.dto.OvertimeRecordDto;
import com.aact.overtime.service.OvertimeRecordService;
import lombok.RequiredArgsConstructor;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.util.List;

/**
 * 시간외근무 개인별 내역 API
 * Base URL: /api/overtime/records
 */
@RestController
@RequestMapping("/api/overtime/records")
@RequiredArgsConstructor
public class OvertimeRecordController {

    private final OvertimeRecordService overtimeRecordService;

    /** POST /api/overtime/records — 단건 등록 */
    @PostMapping
    public ResponseEntity<OvertimeRecordDto.Response> create(
            @RequestBody OvertimeRecordDto.Request request) {
        return ResponseEntity.ok(overtimeRecordService.create(request));
    }

    /** POST /api/overtime/records/batch — 일괄 등록 */
    @PostMapping("/batch")
    public ResponseEntity<List<OvertimeRecordDto.Response>> createAll(
            @RequestBody List<OvertimeRecordDto.Request> requests) {
        return ResponseEntity.ok(overtimeRecordService.createAll(requests));
    }

    /** GET /api/overtime/records/{id} — 단건 조회 */
    @GetMapping("/{id}")
    public ResponseEntity<OvertimeRecordDto.Response> getById(@PathVariable Long id) {
        return ResponseEntity.ok(overtimeRecordService.getById(id));
    }

    /**
     * GET /api/overtime/records?yearMonth=2026-04
     * GET /api/overtime/records?yearMonth=2026-04&department=조업부
     * GET /api/overtime/records?yearMonth=2026-04&employeeName=김보라
     */
    @GetMapping
    public ResponseEntity<List<OvertimeRecordDto.Response>> getList(
            @RequestParam String yearMonth,
            @RequestParam(required = false) String department,
            @RequestParam(required = false) String employeeName) {

        if (department != null) {
            return ResponseEntity.ok(overtimeRecordService.getByYearMonthAndDepartment(yearMonth, department));
        }
        if (employeeName != null) {
            return ResponseEntity.ok(overtimeRecordService.getByYearMonthAndEmployee(yearMonth, employeeName));
        }
        return ResponseEntity.ok(overtimeRecordService.getByYearMonth(yearMonth));
    }

    /** PUT /api/overtime/records/{id} — 수정 */
    @PutMapping("/{id}")
    public ResponseEntity<OvertimeRecordDto.Response> update(
            @PathVariable Long id,
            @RequestBody OvertimeRecordDto.Request request) {
        return ResponseEntity.ok(overtimeRecordService.update(id, request));
    }

    /** DELETE /api/overtime/records/{id} — 삭제 */
    @DeleteMapping("/{id}")
    public ResponseEntity<Void> delete(@PathVariable Long id) {
        overtimeRecordService.delete(id);
        return ResponseEntity.noContent().build();
    }
}
