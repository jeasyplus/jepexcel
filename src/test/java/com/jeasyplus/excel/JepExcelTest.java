package com.jeasyplus.excel;


import com.alibaba.excel.converters.url.UrlImageConverter;
import com.jeasyplus.excel.converter.LocalDateConverter;
import com.jeasyplus.excel.io.FileLoader;
import com.jeasyplus.excel.io.FileWriter;
import com.jeasyplus.excel.reader.CellReader;
import com.jeasyplus.excel.reader.ExcelFile;
import com.jeasyplus.excel.replace.ExcelReplace;
import com.jeasyplus.excel.template.DataContext;
import com.jeasyplus.excel.template.DataConverter;
import com.jeasyplus.excel.template.ExcelTemplate;
import com.jeasyplus.excel.entity.Book;
import com.jeasyplus.excel.entity.TableObject;
import com.jeasyplus.excel.writer.ExcelObj;



import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class JepExcelTest {

    /**
     * 读取excel数据
     * @throws IOException
     */
    public void readData() throws IOException {
        ExcelFile.row("/Users/jepExcel/1654157465882063.xlsx", row->{
            if(row == null){
                return;
            }
            int rowNum = row.getRowNum();
            System.out.println(rowNum);
            System.out.println(CellReader.value(row,2));
            System.out.println(CellReader.toString(row,2));
        },6);
    }

    /**
     * 读取excel数据，可随时结束读取
     * @throws IOException
     */
    public void readLine() throws IOException {
        ExcelFile.next("/Users/jepExcel/1654157465882063.xlsx",1, row->{
            System.out.println(CellReader.value(row,2));
            System.out.println(CellReader.toString(row,2));
            //当返回false时，停止读取数据
            return row != null;
        },8);
    }

    /**
     * 使用数组创建excel，适用不规律的数据
     * @throws IOException
     */
    public void createData() throws IOException {
        //创建excel对象
        ExcelObj excelObj = ExcelObj.create("/Users/jepExcel/list.xlsx");
        //指定sheet
        excelObj.sheet("测试sheet");
        //设置每行数据
        excelObj.row(Arrays.asList("lee",18,"歌手"));
        excelObj.row(Arrays.asList("rain",28,"歌手"));
        excelObj.row(Arrays.asList("seven",28,"演员"));
        //指定sheet
        excelObj.sheet("测试sheet2");
        excelObj.row(Arrays.asList("姓名","年龄","职业","描述"));
        excelObj.row(Arrays.asList("lee",28,"艺术家","画画"));
        excelObj.row(Arrays.asList("rain",38,"工人","生成产品"));
        excelObj.row(Arrays.asList("seven",38,"农民","种庄稼"));

        excelObj.sheet("测试sheet");
        excelObj.row(Arrays.asList("seven",38,"农民","也参与种庄稼"));
        excelObj.row(Arrays.asList("rain",38,"工人","也参与生成产品"));

        excelObj.sheet("测试sheet2");
        excelObj.row(Arrays.asList("lee",18,"也参与画画"));

        //生成excel
        excelObj.writer();
    }

    /**
     * 使用模板结合数生成excel
     * @throws IOException
     */
    private void createByTemplate() throws IOException {
        List<Book> bookList = new ArrayList<>();
        bookList.add(new Book("被讨厌的勇气",300));
        bookList.add(new Book("金字塔原理",500));

        List<Book> bookList2 = new ArrayList<>();
        bookList2.add(new Book("原则",350));
        bookList2.add(new Book("高效能人的7个习惯",800));

        ExcelTemplate.fill(
                //将两组数据放入不同的sheet中
                new DataContext(0,bookList).add(1,bookList2),
                //模板
                FileLoader.getInputStream("/Users/jepExcel/book_tpl.xlsx"),
                //生成的excel
                FileWriter.getOutStream("/Users/jepExcel/book.xlsx"));
    }

    /**
     * 注册数据类型转换器
     * @throws IOException
     */
    private void dataConverter() throws IOException {
        List<Book> bookList = new ArrayList<>();
        bookList.add(new Book("被讨厌的勇气",300));
        bookList.add(new Book("金字塔原理",500));
        List<Book> bookList2 = new ArrayList<>();
        bookList2.add(new Book("原则",350));
        bookList2.add(new Book("高效能人的7个习惯",800));
        int sheetAt = 1;
        ExcelTemplate.fill(
                new DataContext(0,bookList).add(sheetAt,bookList2),
                //注册两个数据类型转换器
                DataConverter.create()
                        .add(new LocalDateConverter())
                        .add(new UrlImageConverter()),

                FileLoader.getInputStream("/Users/jepExcel/book_tpl.xlsx"),
                FileWriter.getOutStream("/Users/jepExcel/book.xlsx"));
    }

    /**
     * 使用不规则的模板，生成excel,适用报表之类的场景
     * @throws IOException
     */
    private void replace() throws IOException {
        //replace方法重载了多个参数，如：是否需要关闭输出流，默认：关闭
        //最多处理完多少行后就结束
        int maxRowNum = 12;
        ExcelReplace.replace(
                FileLoader.getInputStream("/Users/jepExcel/template.xlsx"),
                FileWriter.getOutStream("/Users/jepExcel/report.xlsx"),
                //使用一个对象存放数据
                new TableObject("Lee",28,"歌手","非常好"),
                //指定sheet
                0
                //如果不能正确读取完整的行，可使用maxRowNum参数强制读完指定行数，默认可以不使用
                ,12);
    }

}
