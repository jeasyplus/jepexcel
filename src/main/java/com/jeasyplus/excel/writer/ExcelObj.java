package com.jeasyplus.excel.writer;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class ExcelObj {

    private String path;

    private int sheetAt;

    private Workbook workbook;

    private Sheet sheet;

    private int rowNum;

    public static ExcelObj create(String path) throws FileNotFoundException {
       return new ExcelObj(path);
    }


    public ExcelObj(String path) throws FileNotFoundException {
        if (path == null) {
            throw new RuntimeException("File path parameter cannot be empty");
        }
        this.path = path;
        workbook = new XSSFWorkbook();
        sheetAt = -1;
    }

    public ExcelObj sheet(String sheetName) {
        sheet = workbook.createSheet(sheetName);
        rowNum = 0;
        sheetAt++;
        return this;
    }

    public void row(List rowData) {
        Row row = sheet.createRow(rowNum++);
        if (rowData == null || rowData.isEmpty()) {
            return;
        }
        int cellIndex = 0;
        for (Object data : rowData) {
            CellWriter.setValue(row, cellIndex++, data);
        }
    }

    public void writer() throws IOException {
        try(FileOutputStream out = new FileOutputStream(path)) {
            workbook.write(out);
        }
    }

}
