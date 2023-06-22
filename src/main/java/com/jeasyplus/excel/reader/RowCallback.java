package com.jeasyplus.excel.reader;

import org.apache.poi.ss.usermodel.Row;

public interface RowCallback {

    public void call(Row row);

}
