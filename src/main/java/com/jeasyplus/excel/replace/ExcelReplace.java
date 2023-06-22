package com.jeasyplus.excel.replace;

import com.jeasyplus.excel.reader.CellReader;
import org.apache.poi.ss.usermodel.*;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import java.util.Map;

public class ExcelReplace {


    public static void replace(InputStream templateStream, OutputStream excelStream, Object obj, int sheetAt) throws IOException {
        replace(templateStream,excelStream,obj,sheetAt,true,-1);
    }

    public static void replace(InputStream templateStream, OutputStream excelStream, Object obj, int sheetAt,int maxRowNum) throws IOException {
        replace(templateStream,excelStream,obj,sheetAt,true,maxRowNum);
    }
    public static void replace(InputStream templateStream, OutputStream excelStream, Object obj, int sheetAt,boolean autoCloseable,int maxRowNum) throws IOException {
        try {
            Workbook workbook = WorkbookFactory.create(templateStream);
            if (obj != null) {
                Map<String, Object> dataMap = CellReplaceObject.toMap(obj);
                if (!dataMap.isEmpty()) {
                    Sheet sheet = workbook.getSheetAt(sheetAt);
                    int firstRowNum = sheet.getFirstRowNum();
                    int lastRowNum = sheet.getLastRowNum();
                    if(maxRowNum > 0){
                        lastRowNum = maxRowNum;
                    }
                    for (int rowNum = firstRowNum; rowNum < lastRowNum; rowNum++) {
                        Row row = sheet.getRow(rowNum);
                        if (row != null) {
                            short lastCellNum = row.getLastCellNum();
                            for (int cellRow = 0; cellRow < lastCellNum; cellRow++) {
                                Cell cell = row.getCell(cellRow);
                                if (cell != null) {
                                    String value = CellReader.toString(cell);
                                    CellReplacement.replaceCell(cell, value, dataMap);
                                }
                            }
                        }
                    }
                }
            }
            workbook.write(excelStream);
            excelStream.flush();
        }finally {
            if(autoCloseable){
                excelStream.close();
            }
        }
    }

}
