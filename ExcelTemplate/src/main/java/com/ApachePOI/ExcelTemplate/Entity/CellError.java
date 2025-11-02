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

    public int getRowIndex() {
        return rowIndex;
    }

    public void setRowIndex(int rowIndex) {
        this.rowIndex = rowIndex;
    }

    public int getColumnIndex() {
        return columnIndex;
    }

    public void setColumnIndex(int columnIndex) {
        this.columnIndex = columnIndex;
    }

    public String getErrorMessage() {
        return errorMessage;
    }

    public void setErrorMessage(String errorMessage) {
        this.errorMessage = errorMessage;
    }

}
