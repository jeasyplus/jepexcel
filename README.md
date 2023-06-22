# jepExcel - JavaEasyPlusExcel

jepExcel 是一个用于简化对 Apache POI 和 Alibaba easyExcel 的复杂 API 进行包装的 Java Excel 操作项目。它旨在方便处理一些需求简单的场景。

## 功能特性

- 读取 Excel 数据
- 逐行读取 Excel 数据
- 使用数组创建 Excel，适用于不规则的数据
- 使用模板结合数据生成 Excel
- 使用不规则的模板生成 Excel，适用于报表等场景

## 安装和部署
请根据下列步骤进行安装和部署：
``` xml
<dependency>
    <groupId>com.jeasyplus</groupId>
    <artifactId>jepexcel</artifactId>
    <version>1.0.0</version>
</dependency>
```

## 示例

### 读取 Excel 数据

```java
public void readData() throws IOException {
    ExcelFile.row("/Users/jepExcel/data.xlsx", row -> {
        int rowNum = row.getRowNum();
        System.out.println(rowNum);
        System.out.println(CellReader.value(row,2));
        System.out.println(CellReader.toString(row,2));
    });
}
```
### 逐行读取 Excel 数据

```java
public void readLine() throws IOException {
    ExcelFile.next("/Users/jepExcel/data.xlsx", row -> {
        System.out.println(CellReader.value(row,2));
        System.out.println(CellReader.toString(row,2));
        // 当返回 false 时，停止读取数据
        return false;
    });
}


```

### 使用数组创建 Excel

```java
public void createData() throws IOException {
    // 创建 Excel 对象
    ExcelObj excelObj = ExcelObj.create("/Users/jepExcel/list.xlsx");
    // 指定 sheet
    excelObj.sheet("测试sheet");
    // 设置每行数据
    excelObj.row(Arrays.asList("lee",18,"歌手"));
    excelObj.row(Arrays.asList("rain",28,"歌手"));
    excelObj.row(Arrays.asList("seven",28,"演员"));
    // 指定 sheet
    excelObj.sheet("测试sheet2");
    excelObj.row(Arrays.asList("姓名","年龄","职业","描述"));
    excelObj.row(Arrays.asList("lee",28,"艺术家","画画"));
    excelObj.row(Arrays.asList("rain",38,"工人","生成产品"));
    excelObj.row(Arrays.asList("seven",38,"农民","种庄稼"));
    // 生成 Excel
    excelObj.writer();
}

```

### 使用模板结合数据生成 Excel

```java
private void createByTemplate() throws IOException {
    List<Book> bookList = new ArrayList<>();
    bookList.add(new Book("被讨厌的勇气",300));
    bookList.add(new Book("金字塔原理",500));

    List<Book> bookList2 = new ArrayList<>();
    bookList2.add(new Book("原则",350));
    bookList2.add(new Book("高效能人的7个习惯",800));

    ExcelTemplate.fill(
            // 将两组数据放入不同的 sheet 中
            new DataContext(0,bookList).add(1,bookList2),
            // 模板
            FileLoader.getInputStream("/Users/jepExcel/book_tpl.xlsx"),
            // 生成的 Excel
            FileWriter.getOutStream("/Users/jepExcel/book.xlsx"));
}

```
### 注册数据类型转换器

```java
private void dataConverter() throws IOException {
    List<Book> bookList = new ArrayList<>();
    bookList.add(new Book("被讨厌的勇气",300));
    bookList.add(new Book("金字塔原理",500));
    List<Book> bookList2 = new ArrayList<>();
    bookList2.add(new Book("原则",350));
    bookList2.add(new Book("高效能人的7个习惯",800));

    ExcelTemplate.fill(
            new DataContext(0,bookList).add(1,bookList2),
            // 注册两个数据类型转换器
            DataConverter.create()
                    .add(new LocalDateConverter())
                    .add(new UrlImageConverter()),

            FileLoader.getInputStream("/Users/jepExcel/book_tpl.xlsx"),
            FileWriter.getOutStream("/Users/jepExcel/book.xlsx"));
}

```

### 使用不规则的模板生成 Excel
![示例图片](https://jeasyplus.com/images/ref/jepexcel/example.png)

```java
private void replace() throws IOException {
    // replace 方法重载了多个参数，如：是否需要关闭输出流，默认：关闭
    ExcelReplace.replace(
            FileLoader.getInputStream("/Users/jepExcel/template.xlsx"),
            FileWriter.getOutStream("/Users/jepExcel/report.xlsx"),
            // 使用一个对象存放数据
            new TableObject("Lee",28,"歌手","非常好"),
            // 指定 sheet
            0
            // 如果不能正确读取完整的行，可使用 maxRowNum 参数强制读完指定行数，默认可以不使用
            ,12);
}

```

## 技术栈
- Apache POI
- Alibaba EasyExcel

## 贡献指南
如果您希望为项目做出贡献，请阅读贡献指南并提交问题、建议或 Pull 请求。

**版权和许可**
该项目使用 MIT 许可证。详细信息请参阅 LICENSE 文件。

## 作者
作者：yonghui
邮箱：xyhui321@gmail.com


## 常见问题（FAQ）
这里列出一些常见问题和解答，以帮助用户解决一些常见的疑问。

如果您有其他问题或需要更多信息，请随时联系我们。



