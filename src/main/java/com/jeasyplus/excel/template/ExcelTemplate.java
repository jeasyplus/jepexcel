package com.jeasyplus.excel.template;

import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.write.builder.ExcelWriterSheetBuilder;
import com.alibaba.excel.write.metadata.WriteSheet;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

public class ExcelTemplate {

    public static void fill(DataContext dataContext,  InputStream excelTemplate, OutputStream excelFile) throws IOException {
        fill(dataContext, DataConverter.create(),excelTemplate,excelFile);
    }

    public static void fill(DataContext dataContext, DataConverter converter, InputStream excelTemplate, OutputStream excelFile) throws IOException {
        try {
            ExcelWriter excelWriter = EasyExcelFactory.write(excelFile).withTemplate(excelTemplate).build();
            dataContext.getSheetList().forEach(sheet -> {
                createSheet(sheet.getSheet(), sheet.getData(), converter.list(), excelWriter);
            });
            excelWriter.finish();
        }finally {
            excelFile.close();
        }
    }

    private static void createSheet(int sheetNo, Object data, List<Converter> converterList, ExcelWriter excelWriter) {
        ExcelWriterSheetBuilder builder = EasyExcelFactory.writerSheet(sheetNo);
        if (converterList != null) {
            converterList.forEach(converter -> {
                builder.registerConverter(converter);
            });
        }
        WriteSheet sheet = builder.build();
        excelWriter.fill(data, sheet);
    }

}
