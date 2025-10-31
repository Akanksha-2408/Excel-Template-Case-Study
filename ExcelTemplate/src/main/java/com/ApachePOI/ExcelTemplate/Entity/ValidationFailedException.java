package com.ApachePOI.ExcelTemplate.Entity;

import org.apache.poi.ss.usermodel.Workbook;

public class ValidationFailedException extends RuntimeException {
    private final Workbook errorWorkbook;

    public ValidationFailedException(String message, Workbook errorWorkbook) {
        super(message);
        this.errorWorkbook = errorWorkbook;
    }

    public Workbook getErrorWorkbook() {
        return errorWorkbook;
    }
}
