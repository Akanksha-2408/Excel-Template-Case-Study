package com.ApachePOI.ExcelTemplate.Service;

import com.ApachePOI.ExcelTemplate.Entity.CellError;
import com.ApachePOI.ExcelTemplate.Entity.Employee;
import com.ApachePOI.ExcelTemplate.Repository.EmployeeRepo;
import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

@Service
public class ServiceClass {

    @Autowired
    EmployeeRepo employeeRepo;

    private static final int ROW_LIMIT = 7;
    private static final String[] headers = {"ID", "Name", "Age", "Salary", "DOB"};
    private static final String[] EXPECTED_HEADERS = {"ID", "Name", "Age", "Salary", "DOB"};

    private List<CellError> validationErrors;
    private List<Employee> validEmployees;

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
        int columnCount = 0;
        for(String header : headers) {
            Cell cell = headerRow.createCell(columnCount++);
            cell.setCellValue(header);
            cell.setCellStyle(headerStyle);
        }


        //5. Applying data validation Rules
        DataValidationHelper validationHelper = sheet.getDataValidationHelper();

        //1. Column 0/A
        CellRangeAddressList numericRegion = new CellRangeAddressList(1, ROW_LIMIT, 0, 0);  // Select Column

        DataValidationConstraint idConstraint = validationHelper.createNumericConstraint(  // Select contraint type
                DataValidationConstraint.ValidationType.INTEGER,
                DataValidationConstraint.OperatorType.BETWEEN,
                "100",
                "150"
        );

        DataValidation idValidation = validationHelper.createValidation(idConstraint, numericRegion);  // Set Validation

        // Error Handling
        idValidation.setShowErrorBox(true);
        idValidation.setErrorStyle(DataValidation.ErrorStyle.STOP);
        idValidation.createErrorBox("Wrong ID", "ID must be a number");

        sheet.addValidationData(idValidation);

        //2. Column 1/B

        String cellRef = "B2";
        String noDigitsFormula = "SUMPRODUCT(ISNUMBER(MID(" + cellRef + ", ROW(INDIRECT(\"1:\"&LEN(" + cellRef + "))), 1)*1)*1)=0";
        String specialChars = "{\"!\", \"@\", \"#\", \"$\", \"%\", \"^\", \"&\", \"*\", \"(\", \")\", \"-\", \"+\", \"=\", \"\"\"\", \";\", \":\", \"<\", \">\", \",\", \".\", \"/\", \"\\\\\", \"|\", \" \"}";
        String noSpecialCharsFormula = "SUMPRODUCT(--ISERROR(FIND(" + specialChars + ", " + cellRef + ")))=24";
        String finalFormula = "AND(" + noDigitsFormula + ", " + noSpecialCharsFormula + ", LEN(" + cellRef + ")>=2, LEN(" + cellRef + ")<=30)";

        CellRangeAddressList stringRegion1 = new CellRangeAddressList(1, ROW_LIMIT, 1, 1);  // Select Column
        DataValidationConstraint nameConstraint = validationHelper.createCustomConstraint(  // Select type
                finalFormula
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
        CellRangeAddressList dateRegion = new CellRangeAddressList(1, ROW_LIMIT, 4, 4);
        DataValidationConstraint dobConstraint = validationHelper.createDateConstraint(
                DataValidationConstraint.OperatorType.GREATER_OR_EQUAL,
                "2025/01/01",
                "DATE",
                null
        );

        DataValidation dobValidation = validationHelper.createValidation(dobConstraint, dateRegion);

        dobValidation.setShowErrorBox(true);
        dobValidation.setErrorStyle(DataValidation.ErrorStyle.STOP);
        dobValidation.createErrorBox("Wrong DOB", "DOB must be in YYYY-MM-DD format");

        sheet.addValidationData(dobValidation);

        workbook.write(response.getOutputStream());
        workbook.close();
    }

//    public void uploadFile(InputStream file) throws IOException {
//
//        List<Employee> employees = new ArrayList<Employee>();
//        final String[] EXPECTED_HEADERS = {"ID", "Name", "age", "salary", "DOB"};
//
//        try (Workbook workbook = WorkbookFactory.create(file)) {
//            Sheet sheet = workbook.getSheetAt(0);
//
//            //Save excel data in database
//            if(validateHeaders(file) && validateData(file)) {
//                sheet.forEach(row -> {
//
//                });
//            } else {
//                generateErrorFile(file);
//            }
//            employeeRepo.saveAll(employees);
//        }
//    }

    public void uploadFile(InputStream file) throws IOException {

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        file.transferTo(baos);
        byte[] fileBytes = baos.toByteArray();

        boolean dataIsValid = validateData(new ByteArrayInputStream(fileBytes));
        boolean headersAreValid = validateHeaders(new ByteArrayInputStream(fileBytes));

        if(dataIsValid && headersAreValid) {
            employeeRepo.saveAll(validEmployees);
        } else {
            generateErrorFile(new ByteArrayInputStream(fileBytes));
        }
    }

    public boolean validateHeaders(InputStream file) throws IOException {

        boolean flag = true;

        try (Workbook workbook = WorkbookFactory.create(file)) {
            Sheet sheet = workbook.getSheetAt(0);

            //Header Validation
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) {
                throw new IOException("Uploaded sheet is empty");
            }
            for (int i = 0; i < EXPECTED_HEADERS.length; i++) {
                Cell headerCell = headerRow.getCell(i);

                if (headerCell == null ||
                        headerCell.getCellType() != CellType.STRING ||
                        !headerCell.getStringCellValue().trim().equalsIgnoreCase(EXPECTED_HEADERS[i])) {

                    String actualHeader = (headerCell == null) ? "BLANK/MISSING" : headerCell.getStringCellValue();
                    flag = false;
                    throw new IllegalArgumentException(
                            "Header mismatch! Expected '" + EXPECTED_HEADERS[i] +
                                    "' at column " + (i + 1) + ", but found '" + actualHeader + "'. Please use the correct template.");

                }
            }
            System.out.println("Headers Validated Successfully");
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return flag;
    }

    public boolean validateData(InputStream file) throws IOException {
        this.validationErrors = new ArrayList();
        this.validEmployees = new ArrayList<>();

        try (Workbook workbook = WorkbookFactory.create(file)) {
            Sheet sheet = workbook.getSheetAt(0);
            sheet.forEach(row -> {
                if(row.getRowNum() == 0) {
                    return;
                }
                Employee employee = new Employee();
                boolean rowIsValid = true;
                int rowIndex = row.getRowNum();

                //ID Validation
                Cell idcell = row.getCell(0);
                if(idcell != null && idcell.getCellType() != CellType.NUMERIC) {
                    employee.setId((int) idcell.getNumericCellValue());
                } else {
                    String errorMsg = (idcell == null || idcell.getCellType() == CellType.BLANK) ? "ID is Empty": "ID must be a number";
                    validationErrors.add(new CellError(rowIndex, 0, errorMsg));
                    rowIsValid = false;
                }

                //Name Validation
                Cell namecell = row.getCell(1);
                if(namecell != null && namecell.getCellType() == CellType.FORMULA) {
                    employee.setName(namecell.getStringCellValue());
                } else {
                    validationErrors.add(new CellError(rowIndex, 1, "Name must be String"));
                    rowIsValid = false;
                }

                //Age Validation
                Cell agecell = row.getCell(2);
                if(agecell != null && agecell.getCellType() == CellType.NUMERIC) {
                    employee.setAge((int) agecell.getNumericCellValue());
                } else {
                    validationErrors.add(new CellError(rowIndex, 2, "Age must be a number"));
                    rowIsValid = false;
                }

                //Salary Validation
                Cell salarycell = row.getCell(3);
                if(salarycell != null && salarycell.getCellType() == CellType.NUMERIC) {
                    employee.setSalary(salarycell.getNumericCellValue());
                } else {
                    validationErrors.add(new CellError(rowIndex, 3, "Salary must be a number"));
                    rowIsValid = false;
                }

                //DOB Validation
                Cell dobcell = row.getCell(4);
                if(dobcell != null && dobcell.getCellType() == CellType.NUMERIC) {
                    employee.setDOB(dobcell.getDateCellValue());
                } else {
                    validationErrors.add(new CellError(rowIndex, 4, "DOB must be a valid date format."));
                    rowIsValid = false;
                }

                // Add to valid list if the whole row passed
                if (rowIsValid) {
                    validEmployees.add(employee);
                }

            });
        }
        // Return true if the error list is empty
        return validationErrors.isEmpty();
    }

//    public Workbook generateErrorFile(InputStream file) throws IOException {
    public void generateErrorFile(InputStream file) throws IOException {

        // 1. Re-read the original file uploaded by the user
        Workbook errorWorkbook;

        try {
            errorWorkbook = WorkbookFactory.create(file);

        } catch (Exception e) {

            // Fallback if file is corrupted, create an empty workbook
            errorWorkbook = new XSSFWorkbook();
            Sheet errorSheet = errorWorkbook.createSheet("Error Report");
            Row headerRow = errorSheet.createRow(0);
            for(int i = 0; i < EXPECTED_HEADERS.length; i++) {
                headerRow.createCell(i).setCellValue(EXPECTED_HEADERS[i]);
            }
//            return errorWorkbook;
        }

        Sheet errorSheet = errorWorkbook.getSheetAt(0);
        Drawing<?> drawing = errorSheet.createDrawingPatriarch();

        // 2. Define Cell Styles for Errors (Red Highlight)
        CellStyle errorStyle = errorWorkbook.createCellStyle();
        errorStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
        errorStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // 3. Apply Errors to the Sheet
        for (CellError error : validationErrors) {
            Row row = errorSheet.getRow(error.rowIndex);
            if (row == null) continue;

            Cell cell = row.getCell(error.columnIndex);
            if (cell == null) {
                cell = row.createCell(error.columnIndex); // Create if null
            }

            // a. Apply the Red Highlight
            cell.setCellStyle(errorStyle);

            // b. Add Comments
            // Create a comment box that spans from column C to D, row 2 to 3
            ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, error.columnIndex + 1, error.rowIndex, error.columnIndex + 3, error.rowIndex + 4);
            Comment comment = drawing.createCellComment(anchor);
            comment.setString(errorWorkbook.getCreationHelper().createRichTextString("ERROR: " + error.errorMessage));
            comment.setAuthor("Data Validator");
            cell.setCellComment(comment);
        }

//        return errorWorkbook;
    }

    public List<Employee> findAll() {
        return employeeRepo.findAll();
    }

}
