package com.jeasyplus.excel.reader;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

import static org.apache.poi.ss.usermodel.CellType.*;

public class CellReader {

    public static String toString(Row row, int cell){
        if(row == null){
            return null;
        }
        return toString(row.getCell(cell));
    }

    public static Number toNumber(Row row,int cell){
        return toNumber(row.getCell(cell));
    }

    public static Boolean toBoolean(Row row,int cell){
        return toBoolean(row.getCell(cell));
    }

    public static Object value(Row row,int cell){
        if(row == null){
            return null;
        }
        return value(row.getCell(cell));
    }

    public static String toString(Cell cell){
        if(cell == null){
            return null;
        }
        CellType cellType = cell.getCellType();
        if(STRING.equals(cellType)){
            return cell.getStringCellValue();
        }
        if(FORMULA.equals(cellType)){
            return cell.getCellFormula();
        }
        Object value = value(cell);
        return String.valueOf(value);
    }

    public static Number toNumber(Cell cell){
        if(cell == null){
            return null;
        }
        CellType cellType = cell.getCellType();
        if(NUMERIC.equals(cellType)) {
            return cell.getNumericCellValue();
        }
        Object value = value(cell);
        return Double.valueOf(String.valueOf(value));
    }

    public static Boolean toBoolean(Cell cell){
        if(cell == null){
            return null;
        }
        CellType cellType = cell.getCellType();
        if(BOOLEAN.equals(cellType)){
            return cell.getBooleanCellValue();
        }
        Object value = value(cell);
        return Boolean.valueOf(String.valueOf(value));
    }

    public static Object value(Cell cell){
        if(cell == null){
            return null;
        }
        CellType cellType = cell.getCellType();
        if(STRING.equals(cellType)){
            return cell.getStringCellValue();
        }
        if(NUMERIC.equals(cellType)) {
            return cell.getNumericCellValue();
        }
        if(FORMULA.equals(cellType)){
            return cell.getCellFormula();
        }
        if(BOOLEAN.equals(cellType)){
            return cell.getBooleanCellValue();
        }
        if(BLANK.equals(cellType)){
            return cell.toString();
        }
        return cell.toString();
    }
}
