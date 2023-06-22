package com.jeasyplus.excel.replace;

import com.alibaba.excel.util.StringUtils;
import com.jeasyplus.excel.writer.CellWriter;
import org.apache.poi.ss.usermodel.Cell;

import java.util.Map;

public class CellReplacement {

    private static final String FILL_PREFIX = "{";
    private static final String FILL_SUFFIX = "}";
    private static final String COLLECTION_PREFIX = ".";


    public static void replaceCell(Cell cell, String value, Map<String, Object> dataMap) {
        String cellValue = replace(value, 0, dataMap);
        if(!value.equals(cellValue)){
            CellWriter.setValue(cell, cellValue);
        }
    }

    private static String replace(String value, int start, Map<String, Object> dataMap) {
        for (int i = start; i < value.length(); i++) {
            int prefix = value.indexOf(FILL_PREFIX + COLLECTION_PREFIX, start);
            int suffix = value.indexOf(FILL_SUFFIX, start);
            if (prefix == -1 || suffix == -1 || prefix > suffix) {
                break;
            }
            String prop = value.substring(prefix + 2,suffix);
            Object cellValue = dataMap.get(prop);
            cellValue = cellValue == null ? "" : cellValue;
            value = value.replace(FILL_PREFIX + COLLECTION_PREFIX + prop + FILL_SUFFIX, String.valueOf(cellValue));
            start = prefix + i;
        }
        return value;
    }

}
