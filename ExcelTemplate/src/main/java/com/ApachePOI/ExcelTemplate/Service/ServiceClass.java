package com.ApachePOI.ExcelTemplate.Service;

import com.ApachePOI.ExcelTemplate.Entity.Employee;
import com.ApachePOI.ExcelTemplate.Repository.EmployeeRepo;
import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import java.io.IOException;
import java.io.InputStream;
import java.util.LinkedList;
import java.util.List;

@Service
public class ServiceClass {

    @Autowired
    EmployeeRepo employeeRepo;

    private static final int ROW_LIMIT = 7;

    public void downloadFile(HttpServletResponse response) throws IOException {

        //1. Set response headers for Excel file download
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");  //Client will get response in this (Excel) format
        response.setHeader("Content-Disposition", "attachment; filename=Template.xlsx");
        // response.setHeader() - adds HTTP header to the response
        // Content-Disposition - This is the header name, which is used to indicate whether the content should be displayed inline (within the browser window) or handled as an attachment (downloaded).
        // attachment - keyword that instructs the browser to download the content rather than trying to display it


        //2. Create Workbook and Sheet
        Workbook workbook = new XSSFWorkbook();  // Software
        Sheet sheet = workbook.createSheet("Template");  // File


        //3. Define a header style (Not Optional)
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setColor(IndexedColors.BLUE.getIndex());

        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFont(headerFont);


        //4. Create header Row
        Row headerRow = sheet.createRow(0);
        String[] headers = {"ID", "Name", "Age", "Salary", "Department", "DOB"};
        int columnCount = 0;
        for(String header : headers) {
            Cell cell = headerRow.createCell(columnCount++);
            cell.setCellValue(header);
            cell.setCellStyle(headerStyle);
        }


        //5. Applying data validation Rules
        DataValidationHelper validationHelper = sheet.getDataValidationHelper();

        //1. Column 0/A
        CellRangeAddressList patternRegion = new CellRangeAddressList(1, ROW_LIMIT, 0, 0);  // Select Column
        String pattern = "ISNUMBER(MATCH(\"???####\", A" + 0 + ", 0))";
        DataValidationConstraint idConstraint = validationHelper.createCustomConstraint(  // Select type
                pattern
        );

        DataValidation idValidation = validationHelper.createValidation(idConstraint, patternRegion);  // Set Validation

        // Error Handling
        idValidation.setShowErrorBox(true);
        idValidation.setErrorStyle(DataValidation.ErrorStyle.STOP);
        idValidation.createErrorBox("Wrong ID", "ID must follow the pattern");

        sheet.addValidationData(idValidation);

        //2. Column 1/B
        CellRangeAddressList stringRegion1 = new CellRangeAddressList(1, ROW_LIMIT, 1, 1);  // Select Column
        DataValidationConstraint nameConstraint = validationHelper.createTextLengthConstraint(  // Select type
                DataValidationConstraint.OperatorType.BETWEEN,
                "2",
                "30"
        );

        DataValidation nameValidation = validationHelper.createValidation(nameConstraint, stringRegion1);  // Set Validation

        //Error Handling
        nameValidation.setShowErrorBox(true);
        nameValidation.setErrorStyle(DataValidation.ErrorStyle.STOP);
        nameValidation.createErrorBox("Wrong Name", "Name must be a string");

        sheet.addValidationData(nameValidation);

        //3. Column 2/C
        CellRangeAddressList numericRegion1 = new CellRangeAddressList(1, ROW_LIMIT, 2, 2);
        DataValidationConstraint ageConstraint = validationHelper.createNumericConstraint(
                DataValidationConstraint.ValidationType.INTEGER,  // what data type validations you are going to provide in next line (range)
                DataValidationConstraint.OperatorType.GREATER_OR_EQUAL,
                "1",
                null
        );

        DataValidation ageValidation = validationHelper.createValidation(ageConstraint, numericRegion1);

        ageValidation.setShowErrorBox(true);
        ageValidation.setErrorStyle(DataValidation.ErrorStyle.STOP);
        ageValidation.createErrorBox("Wrong Age", "Age must be an integer");

        sheet.addValidationData(ageValidation);

        //4. Column 3/D
        CellRangeAddressList numericRegion2 = new CellRangeAddressList(1, ROW_LIMIT, 3, 3);
        DataValidationConstraint salaryConstraint = validationHelper.createNumericConstraint(
                DataValidationConstraint.ValidationType.DECIMAL,  // what data type validations you are going to provide in next line (range)
                DataValidationConstraint.OperatorType.GREATER_OR_EQUAL,
                "10000",
                null
        );

        DataValidation salaryValidation = validationHelper.createValidation(salaryConstraint, numericRegion2);

        salaryValidation.setShowErrorBox(true);
        salaryValidation.setErrorStyle(DataValidation.ErrorStyle.STOP);
        salaryValidation.createErrorBox("Wrong salary", "Salary must be an Number");

        sheet.addValidationData(salaryValidation);

        //Column 4/E
        CellRangeAddressList stringRegion2 = new CellRangeAddressList(1, ROW_LIMIT, 4, 4);
        DataValidationConstraint departmentConstraint = validationHelper.createTextLengthConstraint(  // Select type
                DataValidationConstraint.OperatorType.BETWEEN,
                "2",
                "30"
        );

        DataValidation departmentValidation = validationHelper.createValidation(departmentConstraint, stringRegion2);  // Set Validation

        //Error Handling
        departmentValidation.setShowErrorBox(true);
        departmentValidation.setErrorStyle(DataValidation.ErrorStyle.STOP);
        departmentValidation.createErrorBox("Wrong department", "Department must be a string");

        sheet.addValidationData(departmentValidation);

        //Column 5/F
        CellRangeAddressList dateRegion = new CellRangeAddressList(1, ROW_LIMIT, 5, 5);
        DataValidationConstraint dobConstraint = validationHelper.createDateConstraint(
                DataValidationConstraint.OperatorType.BETWEEN,
                "2025/01/01",
                "2025/10/30",
                "DATE"
        );

        DataValidation dobValidation = validationHelper.createValidation(dobConstraint, dateRegion);

        dobValidation.setShowErrorBox(true);
        dobValidation.setErrorStyle(DataValidation.ErrorStyle.STOP);
        dobValidation.createErrorBox("Wrong DOB", "DOB must be in YYYY-MM-DD format");

        sheet.addValidationData(dobValidation);

        workbook.write(response.getOutputStream());
        workbook.close();
    }

    public void uploadFile(InputStream file) throws IOException {
        List<Employee> employees = new LinkedList<>();

        try (Workbook workbook = WorkbookFactory.create(file)) {
            Sheet sheet = workbook.getSheetAt(0);

            sheet.forEach(row -> {
                Employee employee = new Employee();

                if (row.getRowNum() != 0) {

                    Cell idcell = row.getCell(0);
                    Cell namecell = row.getCell(1);
                    Cell agecell = row.getCell(2);
                    Cell salarycell = row.getCell(3);

                    if(idcell != null && idcell.getCellType() == CellType.STRING) {
                        employee.setId(row.getCell(0).getStringCellValue());
                    } else if(namecell != null && namecell.getCellType() == CellType.BLANK) {
                        System.out.println("Problem in ID: ID is Empty");
                    } else {
                        System.out.println("Problem in ID: ID Does not match the standard pattern");
                    }
                    if(namecell != null && namecell.getCellType() == CellType.STRING) {
                        employee.setName(row.getCell(1).getStringCellValue());
                    } else {
                        System.out.println("Problem in Name");
                    }
                    if(agecell != null && agecell.getCellType() == CellType.NUMERIC) {
                        employee.setAge((int) row.getCell(2).getNumericCellValue());
                    } else {
                        System.out.println("Problem in Age");
                    }
                    if(salarycell != null && salarycell.getCellType() == CellType.NUMERIC) {
                        employee.setSalary(row.getCell(3).getNumericCellValue());
                    } else {
                        System.out.println("Problem in Salary");
                    }

                    employees.add(employee);
                }
            });
            employeeRepo.saveAll(employees);
        }
    }


    public List<Employee> findAll() {
        return employeeRepo.findAll();
    }
}
