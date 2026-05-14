package com.aact.overtime.repository;

import com.aact.overtime.entity.DepartmentSummary;
import org.springframework.data.jpa.repository.JpaRepository;

import java.util.List;
import java.util.Optional;

public interface DepartmentSummaryRepository extends JpaRepository<DepartmentSummary, Long> {

    List<DepartmentSummary> findByApplyYearMonthOrderByDepartmentAsc(String applyYearMonth);

    Optional<DepartmentSummary> findByDepartmentAndApplyYearMonth(String department, String applyYearMonth);
}
