package com.jeasyplus.excel.template;

import java.util.ArrayList;
import java.util.List;

public class DataContext {

    private List<Sheet> sheetList;

    public DataContext(int sheetAt,Object data){
        sheetList = new ArrayList<>();
        add(sheetAt,data);
    }

    public DataContext add(int sheetAt,Object data){
        Sheet sheetObj = new Sheet();
        sheetObj.data=data;
        sheetObj.sheet =sheetAt;
        sheetList.add(sheetObj);
        return this;
    }

    public List<Sheet> getSheetList() {
        return sheetList;
    }

    static class Sheet{
        private int sheet;

        private Object data;

        public Object getData(){
            return data;
        }

        public int getSheet(){
            return sheet;
        }
    }


}
