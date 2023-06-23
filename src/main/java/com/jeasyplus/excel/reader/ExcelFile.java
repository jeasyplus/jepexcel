package com.jeasyplus.excel.reader;


import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;


import java.io.FileInputStream;
import java.io.IOException;

public class ExcelFile {

    public static Workbook workbook(String path) throws IOException {
        if (path == null) {
            throw new RuntimeException("File path parameter cannot be empty");
        }
        FileInputStream in = null;
        try {
            in = new FileInputStream(path);
            Workbook workbook = WorkbookFactory.create(in);
            return workbook;
        } finally {
            if (in != null) {
                in.close();
            }
        }
    }

    public static void next(String path, RowDataCallback callback) throws IOException {
        next(path, callback, 0);
    }


    public static void next(String path, int sheetAt, RowDataCallback callback) throws IOException {
        next(path, sheetAt, callback, 0);
    }

    public static void next(String path, int sheetAt, RowDataCallback callback,int startRowNum) throws IOException {
        next(path, sheetAt, callback, startRowNum,0);
    }

    public static void next(String path, RowDataCallback callback,int startRowNum) throws IOException {
        next(path,  callback,startRowNum,0);
    }

    public static void next(String path, RowDataCallback callback,int startRowNum,int endRowNum) throws IOException {
        next(path, 0, callback,startRowNum,endRowNum);
    }

    public static void next(String path, int sheetAt, RowDataCallback callback, int startRowNum,int endRowNum) throws IOException {
        Workbook workbook = workbook(path);
        Sheet sheet = workbook.getSheetAt(sheetAt);
        int firstRowNum = sheet.getFirstRowNum();
        int lastRowNum = sheet.getLastRowNum();
        if(firstRowNum == lastRowNum){
            throw new RuntimeException("sheet is empty");
        }
        if (startRowNum > 0) {
            firstRowNum = startRowNum;
        }
        if(endRowNum > 0){
            lastRowNum = endRowNum;
        }
        for (int rowIndex = firstRowNum; rowIndex < lastRowNum; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            Boolean next = callback.call(row);
            if (next != null && !next) {
                break;
            }
        }
    }

    public static void row(String path, RowCallback callback) throws IOException {
        row(path, 0, callback);
    }

    public static void row(String path, int sheetAt, final RowCallback callback) throws IOException {
        next(path, sheetAt, row -> {
            callback.call(row);
            return true;
        });
    }

    public static void row(String path, final RowCallback callback, int startRowNum) throws IOException {
      row(path,0,callback,startRowNum);
    }

    public static void row(String path, int sheetAt, final RowCallback callback, int startRowNum) throws IOException {
        row(path,sheetAt,callback,startRowNum,-1);
    }

    public static void row(String path, int sheetAt, final RowCallback callback, int startRowNum,int endRowNum) throws IOException {
        next(path, sheetAt, row -> {
            callback.call(row);
            return true;
        }, startRowNum,endRowNum);
    }

}
