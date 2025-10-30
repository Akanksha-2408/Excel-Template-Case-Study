package com.ApachePOI.ExcelTemplate.Repository;

import com.ApachePOI.ExcelTemplate.Entity.Employee;
import org.springframework.data.jpa.repository.JpaRepository;

public interface EmployeeRepo extends JpaRepository<Employee,Integer> {
}
