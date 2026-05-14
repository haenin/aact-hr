package com.aact.overtime.repository;

import com.aact.overtime.entity.OvertimeRecord;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import java.util.List;

public interface OvertimeRecordRepository extends JpaRepository<OvertimeRecord, Long> {

    /** 특정 연월의 전체 개인별 내역 */
    List<OvertimeRecord> findByApplyYearMonthOrderByDepartmentAscEmployeeNameAsc(String applyYearMonth);

    /** 특정 연월 + 부서 */
    List<OvertimeRecord> findByApplyYearMonthAndDepartmentOrderByWorkDateAsc(
            String applyYearMonth, String department);

    /** 특정 연월 + 직원 */
    List<OvertimeRecord> findByApplyYearMonthAndEmployeeNameOrderByWorkDateAsc(
            String applyYearMonth, String employeeName);

    /** 부서별 시간 합산 (집계표 자동 계산용) */
    @Query("""
        SELECT r.department,
               SUM(r.extensionHours),
               SUM(r.nightHours),
               SUM(r.holidayHours),
               SUM(r.holidayExtensionHours)
        FROM OvertimeRecord r
        WHERE r.applyYearMonth = :yearMonth
        GROUP BY r.department
    """)
    List<Object[]> aggregateByDepartment(@Param("yearMonth") String yearMonth);

    /** 직원별 소계 */
    @Query("""
        SELECT r.employeeName,
               SUM(r.extensionHours),
               SUM(r.nightHours),
               SUM(r.holidayHours),
               SUM(r.holidayExtensionHours)
        FROM OvertimeRecord r
        WHERE r.applyYearMonth = :yearMonth AND r.department = :department
        GROUP BY r.employeeName
    """)
    List<Object[]> aggregateByEmployee(@Param("yearMonth") String yearMonth,
                                       @Param("department") String department);
}
