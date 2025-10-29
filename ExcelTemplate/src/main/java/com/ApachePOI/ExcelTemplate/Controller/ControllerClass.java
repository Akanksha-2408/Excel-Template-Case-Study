package com.ApachePOI.ExcelTemplate.Controller;

import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;

@RestController
public class ControllerClass {

    private static final int ROW_LIMIT = 7;

    @GetMapping("/download")
    public void downloadFile(HttpServletResponse response) throws IOException {

        //1. Set response headers for Excel file download
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");  //Client will get response in this (Excel) format
        response.setHeader("Content-Disposition", "attachment; filename=Template.xltx");
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
        String[] headers = {"ID", "Name", "Department"};
        int columnCount = 0;
        for(String header : headers) {
            Cell cell = headerRow.createCell(columnCount++);
            cell.setCellValue(header);
            cell.setCellStyle(headerStyle);
        }

        //5. Applying data validation Rules to Column 0/A
        CellRangeAddressList compulsoryRegions = new CellRangeAddressList(1,ROW_LIMIT-1, 0, 0); // select rows

        DataValidationHelper validationHelper = sheet.getDataValidationHelper();

        // Compulsory Field
        DataValidationConstraint compulsoryConstraint = validationHelper.createTextLengthConstraint(
                DataValidationConstraint.OperatorType.GREATER_OR_EQUAL,
                "1",
                null
        );

        DataValidation compulsoryValidation = validationHelper.createValidation(
                compulsoryConstraint, compulsoryRegions
        );
        compulsoryValidation.setShowErrorBox(true);
        compulsoryValidation.setErrorStyle(DataValidation.ErrorStyle.STOP);
        compulsoryValidation.createErrorBox("REQUIRED FIELD", "ID Field cannot be empty");
        sheet.addValidationData(compulsoryValidation);

        //6. Applying data validation rules to Column 1/B


        // Auto-size columns for better viewing
        for (int i = 0; i < headers.length; i++) {
            sheet.autoSizeColumn(i);
        }
        
        workbook.write(response.getOutputStream());
        workbook.close();
    }


    @PostMapping("/uploadData")
    public void uploadData(){

    }

}
