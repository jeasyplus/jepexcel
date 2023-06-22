package com.jeasyplus.excel.template;

import com.alibaba.excel.converters.Converter;

import java.util.ArrayList;
import java.util.List;

public class DataConverter {

    private List<Converter> list;

    public static DataConverter create(){
        return new DataConverter();
    }

    public DataConverter add(Converter converter){
        if(list == null){
            list = new ArrayList<>();
        }
        list.add(converter);
        return this;
    }
    public List<Converter> list(){
        return list;
    }

}
