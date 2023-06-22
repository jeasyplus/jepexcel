package com.jeasyplus.excel.writer;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

public class CellWriter {

    public static Cell setValue(Row row,int cellIndex, Object value){
        if(value == null){
            return row.createCell(cellIndex, CellType.BLANK);
        }
        if(value instanceof String){
            Cell cell = row.createCell(cellIndex, CellType.STRING);
            cell.setCellValue((String)value);
            return cell;
        }
        if(value instanceof Number){
            Cell cell = row.createCell(cellIndex, CellType.NUMERIC);
            cell.setCellValue(Double.valueOf(String.valueOf(value)));
            return cell;
        }
        if(value instanceof Boolean){
            Cell cell = row.createCell(cellIndex, CellType.BOOLEAN);
            cell.setCellValue((Boolean)value);
            return cell;
        }
        Cell cell = row.createCell(cellIndex, CellType.STRING);
        cell.setCellValue(String.valueOf(value));
        return cell;
    }

    public static void setValue(Cell cell,Object value){
        if(value == null){
            return;
        }
        if(value instanceof String){
            cell.setCellValue((String)value);
            return;
        }
        if(value instanceof Number){
            cell.setCellValue(Double.valueOf(String.valueOf(value)));
            return;
        }
        if(value instanceof Boolean){
            cell.setCellValue((Boolean)value);
            return;
        }
        cell.setCellValue(String.valueOf(value));
    }
}
