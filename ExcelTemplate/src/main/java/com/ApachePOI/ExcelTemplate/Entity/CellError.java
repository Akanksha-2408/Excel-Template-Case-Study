package com.ApachePOI.ExcelTemplate.Entity;

public class CellError {
    public int rowIndex;
    public int columnIndex;
    public String errorMessage;

    public CellError(int rowIndex, int columnIndex, String errorMessage) {
        this.rowIndex = rowIndex;
        this.columnIndex = columnIndex;
        this.errorMessage = errorMessage;
    }
}
