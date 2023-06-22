package com.jeasyplus.excel.reader;

import org.apache.poi.ss.usermodel.Row;

public interface RowDataCallback {

    public Boolean call(Row row);

}
