package com.jeasyplus.excel.writer;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelObj {

    private String path;

    private int sheetAt;

    private Workbook workbook;

    private SheetMeta sheetMeta;

    private List<SheetMeta> sheetMetaList;


    public static ExcelObj create(String path) throws FileNotFoundException {
        return new ExcelObj(path);
    }


    public ExcelObj(String path) throws FileNotFoundException {
        if (path == null) {
            throw new RuntimeException("File path parameter cannot be empty");
        }
        this.path = path;
        workbook = new XSSFWorkbook();
        sheetAt = 0;
        sheetMetaList = new ArrayList<>();
    }

    public ExcelObj sheet(String sheetName) {
        SheetMeta sheetMetaObj = getSheetMeta(sheetName);
        if(sheetMetaObj != null){
            sheetMeta = sheetMetaObj;
            return this;
        }
        Sheet sheet = workbook.createSheet(sheetName);
        sheetMeta = new SheetMeta(sheetAt++, sheetName, 0, sheet);
        sheetMetaList.add(sheetMeta);
        return this;
    }

    private SheetMeta getSheetMeta(String sheetName) {
        if (sheetName == null || sheetMetaList == null || sheetMetaList.isEmpty()) {
            return null;
        }
        for (SheetMeta meta : sheetMetaList) {
            if(sheetName.equals(meta.getSheetName())){
                return meta;
            }
        }
        return null;
    }

    public void row(List rowData) {
        Sheet sheet = sheetMeta.getSheet();
        Row row = sheet.createRow(sheetMeta.rowNum());
        if (rowData == null || rowData.isEmpty()) {
            return;
        }
        int cellIndex = 0;
        for (Object data : rowData) {
            CellWriter.setValue(row, cellIndex++, data);
        }
    }

    public void writer() throws IOException {
        try (FileOutputStream out = new FileOutputStream(path)) {
            workbook.write(out);
        }
    }

}
