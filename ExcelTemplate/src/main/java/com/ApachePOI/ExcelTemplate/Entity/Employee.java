package com.ApachePOI.ExcelTemplate.Entity;

import jakarta.persistence.*;

import java.util.Date;

@Entity
@Table(name="new-emp")
public class Employee {
    @Id
    @GeneratedValue(strategy=GenerationType.IDENTITY)
    private String id;
    private String name;
    private int age;
    private double salary;
    private String department;
    private Date DOB;

    public Employee() {}

    public Employee(String id, String name, int age, double salary, String department, Date DOB) {
        this.id = id;
        this.name = name;
        this.age = age;
        this.salary = salary;
        this.department = department;
        this.DOB = DOB;
    }

    //getters and setters
    public String getId() {
        return id;
    }
    public void setId(String id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }
    public void setName(String name) {
        this.name = name;
    }

    public int getAge() { return age;}
    public void setAge(int age) { this.age = age;}

    public double getSalary() { return salary;}
    public void setSalary(double salary) {this.salary = salary;}

    public String getDepartment() { return department;}
    public void setDepartment(String department) {this.department = department;}

    public Date getDOB() { return DOB; }
    public void setDOB(Date DOB) { this.DOB = DOB; }

}
