package com.aact.schedule.repository;

import com.aact.schedule.entity.ScheduleRecord;
import org.springframework.data.jpa.repository.JpaRepository;

import java.util.List;
import java.util.Optional;

public interface ScheduleRecordRepository extends JpaRepository<ScheduleRecord, Long> {

    /** 연월 + 부서 전체 (순번 순) */
    List<ScheduleRecord> findByApplyYearMonthAndDepartmentOrderBySeqAsc(
            String applyYearMonth, String department);

    /** 단건 */
    Optional<ScheduleRecord> findByApplyYearMonthAndDepartmentAndEmployeeName(
            String applyYearMonth, String department, String employeeName);

    /** 연월 전체 */
    List<ScheduleRecord> findByApplyYearMonthOrderByDepartmentAscSeqAsc(
            String applyYearMonth);
}
