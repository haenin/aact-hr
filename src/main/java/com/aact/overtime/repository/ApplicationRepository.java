package com.aact.overtime.repository;

import com.aact.overtime.entity.OvertimePayApplication;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import java.util.List;

public interface ApplicationRepository extends JpaRepository<OvertimePayApplication, Long> {

    /** 연월 + 부서 전체 (순번 순) */
    List<OvertimePayApplication> findByApplyYearMonthAndDepartmentOrderByNoAsc(
            String applyYearMonth, String department);

    /** 연월 + 부서 + 성명 */
    List<OvertimePayApplication> findByApplyYearMonthAndDepartmentAndEmployeeNameOrderByNoAsc(
            String applyYearMonth, String department, String employeeName);

    /** 연장/야간/휴일 합계 */
    @Query("""
        SELECT SUM(a.extensionHours), SUM(a.nightHours), SUM(a.holidayHours)
        FROM OvertimePayApplication a
        WHERE a.applyYearMonth = :yearMonth AND a.department = :department
    """)
    List<Object[]> sumByYearMonthAndDepartment(
            @Param("yearMonth") String yearMonth,
            @Param("department") String department);
}
