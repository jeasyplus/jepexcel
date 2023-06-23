package com.jeasyplus.excel.writer;

import org.apache.poi.ss.usermodel.Sheet;

public class SheetMeta {

    private  int sheetAt;
    private String sheetName;
    private int rowNum = 0;
    private Sheet sheet;

    public SheetMeta(int sheetAt, String sheetName, int rowNum, Sheet sheet) {
        this.sheetAt = sheetAt;
        this.sheetName = sheetName;
        this.rowNum = rowNum;
        this.sheet = sheet;
    }

    public int getSheetAt() {
        return sheetAt;
    }

    public void setSheetAt(int sheetAt) {
        this.sheetAt = sheetAt;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public int getRowNum() {
        return rowNum;
    }

    public void setRowNum(int rowNum) {
        this.rowNum = rowNum;
    }

    public Sheet getSheet() {
        return sheet;
    }

    public int rowNum(){
        return rowNum++;
    }

    public void setSheet(Sheet sheet) {
        this.sheet = sheet;
    }
}
