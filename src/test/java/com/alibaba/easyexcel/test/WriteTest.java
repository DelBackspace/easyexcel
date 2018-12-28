package com.alibaba.easyexcel.test;

import cn.hutool.core.date.DateUtil;
import com.alibaba.easyexcel.test.listen.AfterWriteHandlerImpl;
import com.alibaba.easyexcel.test.model.WriteModel;
import com.alibaba.easyexcel.test.util.FileUtil;
import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.metadata.Table;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.util.StringUtils;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.*;

import static com.alibaba.easyexcel.test.util.DataUtil.*;

public class WriteTest {

    private static final String FILE_PATH = "/home/pjz/temple_excel/2007.xlsx";


    private static List<DynamicConfig> DYNAMIC_LIST = new ArrayList<DynamicConfig>();
    private static List<HashMap<String, Object>> DB_RESULT_LIST = new ArrayList<HashMap<String, Object>>();

    private void initConfig(){
        DYNAMIC_LIST.add(new DynamicConfig(1,"字符串列","col_str"));

        DynamicConfig configConvert = new DynamicConfig(1, "整数列转换", "col_int");
        Map<String, String> mapping = new HashMap<String, String>();
        mapping.put("1", "状态一");
        mapping.put("2", "状态二");
        configConvert.setResultMapping(mapping);
        configConvert.setHasMapping(Boolean.TRUE);

        DYNAMIC_LIST.add(configConvert);
        DYNAMIC_LIST.add(new DynamicConfig(1,"金额列","col_decimal"));

        DynamicConfig dateConvert = new DynamicConfig(1,"日期列","col_date");
        dateConvert.setHasDateFormat(Boolean.TRUE);
        dateConvert.setFormatStr("yyyy-MM-dd HH:mm:ss");

        DYNAMIC_LIST.add(dateConvert);
    }

    private void initData(){

        HashMap<String,Object> da1 = new HashMap<String, Object>();
        da1.put("col_str","测试字符串1");
        da1.put("col_int",1);
        da1.put("col_decimal",BigDecimal.valueOf(0.01));
        da1.put("col_date",new Date());
        DB_RESULT_LIST.add(da1);

        HashMap<String,Object> da2 = new HashMap<String, Object>();
        da2.put("col_str","测试字符串2");
        da2.put("col_int",2);
        da2.put("col_decimal",BigDecimal.valueOf(1.00));
        da2.put("col_date",new Date());
        DB_RESULT_LIST.add(da2);

        HashMap<String,Object> da3 = new HashMap<String, Object>();
        da3.put("col_str","测试字符串2");
        da3.put("col_int",3);
        da3.put("col_decimal",BigDecimal.valueOf(1.0123400));
        da3.put("col_date",null);
        DB_RESULT_LIST.add(da3);

    }

    @Test
    public void writeV2007WithCustHead() throws IOException{

        initConfig();
        initData();

        OutputStream out = new FileOutputStream(FILE_PATH);
        ExcelWriter writer = EasyExcelFactory.getWriter(out);
        Sheet sheet1 = new Sheet(1, 0);

        sheet1.setHead(createTestCustHead(DYNAMIC_LIST));
        // 设置自适应宽度
        sheet1.setAutoWidth(Boolean.TRUE);
        writer.write1(createTestCustData(DYNAMIC_LIST, DB_RESULT_LIST), sheet1);
//        writer.write1(createTestCustData(DYNAMIC_LIST, DB_RESULT_LIST), sheet1);
//        writer.write1(createTestCustData(DYNAMIC_LIST, DB_RESULT_LIST), sheet1);

        writer.finish();
        out.close();

    }


    public static List<List<String>> createTestCustHead(List<DynamicConfig> dynamicConfigs){
        //写sheet3  模型上没有注解，表头数据动态传入
        List<List<String>> head = new ArrayList<List<String>>();

        for (DynamicConfig dynamicConfig : dynamicConfigs) {
            head.add(Collections.singletonList(dynamicConfig.getShowHead()));
        }

        return head;
    }

    public static List<List<Object>> createTestCustData(List<DynamicConfig> dynamicConfigs,List<HashMap<String,Object>> dbDataList) {
        List<List<Object>> dataList = new ArrayList<List<Object>>();
        for (HashMap<String, Object> dbDaMap : dbDataList) {
            List<Object> data = new ArrayList<Object>();
            for (DynamicConfig dynamicConfig : dynamicConfigs) {
                if (dynamicConfig.getHasMapping()) {

                    data.add(dynamicConfig.getResultMapping().get(String.valueOf(dbDaMap.get(dynamicConfig.getColName()))));
                    continue;
                }
                if (dynamicConfig.getHasDateFormat()) {
                    data.add(DateUtil.format((Date) dbDaMap.get(dynamicConfig.getColName()),dynamicConfig.getFormatStr()));
                    continue;
                }
                data.add(dbDaMap.get(dynamicConfig.getColName()));

            }
            dataList.add(data);
        }

        return dataList;
    }


    @Test
    public void writeV2007() throws IOException {
        OutputStream out = new FileOutputStream(FILE_PATH);
        ExcelWriter writer = EasyExcelFactory.getWriter(out);
        //写第一个sheet, sheet1  数据全是List<String> 无模型映射关系
        Sheet sheet1 = new Sheet(1, 3);
        sheet1.setSheetName("第一个sheet");

        //设置列宽 设置每列的宽度
        Map columnWidth = new HashMap();
        columnWidth.put(0,10000);columnWidth.put(1,40000);columnWidth.put(2,10000);columnWidth.put(3,10000);
        sheet1.setColumnWidthMap(columnWidth);
        sheet1.setHead(createTestListStringHead());
        //or 设置自适应宽度
        //sheet1.setAutoWidth(Boolean.TRUE);
        writer.write1(createTestListObject(), sheet1);

        //写第二个sheet sheet2  模型上打有表头的注解，合并单元格
        Sheet sheet2 = new Sheet(2, 3, WriteModel.class, "第二个sheet", null);
        sheet2.setTableStyle(createTableStyle());
        //writer.write1(null, sheet2);
        writer.write(createTestListJavaMode(), sheet2);
        //需要合并单元格
        writer.merge(5,20,1,1);

        //写第三个sheet包含多个table情况
        Sheet sheet3 = new Sheet(3, 0);
        sheet3.setSheetName("第三个sheet");
        Table table1 = new Table(1);
        table1.setHead(createTestListStringHead());
        writer.write1(createTestListObject(), sheet3, table1);

        //写sheet2  模型上打有表头的注解
        Table table2 = new Table(2);
        table2.setTableStyle(createTableStyle());
        table2.setClazz(WriteModel.class);
        writer.write(createTestListJavaMode(), sheet3, table2);

        writer.finish();
        out.close();

    }

    @Test
    public void writeV2007WithTemplate() throws IOException {
        InputStream inputStream = FileUtil.getResourcesFileInputStream("temp.xlsx");
        OutputStream out = new FileOutputStream("/Users/jipengfei/2007.xlsx");
        ExcelWriter writer = EasyExcelFactory.getWriterWithTemp(inputStream,out,ExcelTypeEnum.XLSX,true);
        //写第一个sheet, sheet1  数据全是List<String> 无模型映射关系
        Sheet sheet1 = new Sheet(1, 3);
        sheet1.setSheetName("第一个sheet");
        sheet1.setStartRow(20);

        //设置列宽 设置每列的宽度
        Map columnWidth = new HashMap();
        columnWidth.put(0,10000);columnWidth.put(1,40000);columnWidth.put(2,10000);columnWidth.put(3,10000);
        sheet1.setColumnWidthMap(columnWidth);
        sheet1.setHead(createTestListStringHead());
        //or 设置自适应宽度
        //sheet1.setAutoWidth(Boolean.TRUE);
        writer.write1(createTestListObject(), sheet1);

        //写第二个sheet sheet2  模型上打有表头的注解，合并单元格
        Sheet sheet2 = new Sheet(2, 3, WriteModel.class, "第二个sheet", null);
        sheet2.setTableStyle(createTableStyle());
        sheet2.setStartRow(20);
        writer.write(createTestListJavaMode(), sheet2);

        //写第三个sheet包含多个table情况
        Sheet sheet3 = new Sheet(3, 0);
        sheet3.setSheetName("第三个sheet");
        sheet3.setStartRow(30);
        Table table1 = new Table(1);
        table1.setHead(createTestListStringHead());
        writer.write1(createTestListObject(), sheet3, table1);

        //写sheet2  模型上打有表头的注解
        Table table2 = new Table(2);
        table2.setTableStyle(createTableStyle());
        table2.setClazz(WriteModel.class);
        writer.write(createTestListJavaMode(), sheet3, table2);

        writer.finish();
        out.close();

    }

    @Test
    public void writeV2007WithTemplateAndHandler() throws IOException {
        InputStream inputStream = FileUtil.getResourcesFileInputStream("temp.xlsx");
        OutputStream out = new FileOutputStream("/Users/jipengfei/2007.xlsx");
        ExcelWriter writer = EasyExcelFactory.getWriterWithTempAndHandler(inputStream,out,ExcelTypeEnum.XLSX,true,
            new AfterWriteHandlerImpl());
        //写第一个sheet, sheet1  数据全是List<String> 无模型映射关系
        Sheet sheet1 = new Sheet(1, 3);
        sheet1.setSheetName("第一个sheet");
        sheet1.setStartRow(20);

        //设置列宽 设置每列的宽度
        Map columnWidth = new HashMap();
        columnWidth.put(0,10000);columnWidth.put(1,40000);columnWidth.put(2,10000);columnWidth.put(3,10000);
        sheet1.setColumnWidthMap(columnWidth);
        sheet1.setHead(createTestListStringHead());
        //or 设置自适应宽度
        //sheet1.setAutoWidth(Boolean.TRUE);
        writer.write1(createTestListObject(), sheet1);

        //写第二个sheet sheet2  模型上打有表头的注解，合并单元格
        Sheet sheet2 = new Sheet(2, 3, WriteModel.class, "第二个sheet", null);
        sheet2.setTableStyle(createTableStyle());
        sheet2.setStartRow(20);
        writer.write(createTestListJavaMode(), sheet2);

        //写第三个sheet包含多个table情况
        Sheet sheet3 = new Sheet(3, 0);
        sheet3.setSheetName("第三个sheet");
        sheet3.setStartRow(30);
        Table table1 = new Table(1);
        table1.setHead(createTestListStringHead());
        writer.write1(createTestListObject(), sheet3, table1);

        //写sheet2  模型上打有表头的注解
        Table table2 = new Table(2);
        table2.setTableStyle(createTableStyle());
        table2.setClazz(WriteModel.class);
        writer.write(createTestListJavaMode(), sheet3, table2);

        writer.finish();
        out.close();

    }



    @Test
    public void writeV2003() throws IOException {
        OutputStream out = new FileOutputStream("/Users/jipengfei/2003.xls");
        ExcelWriter writer = EasyExcelFactory.getWriter(out, ExcelTypeEnum.XLS,true);
        //写第一个sheet, sheet1  数据全是List<String> 无模型映射关系
        Sheet sheet1 = new Sheet(1, 3);
        sheet1.setSheetName("第一个sheet");

        //设置列宽 设置每列的宽度
        Map columnWidth = new HashMap();
        columnWidth.put(0,10000);columnWidth.put(1,40000);columnWidth.put(2,10000);columnWidth.put(3,10000);
        sheet1.setColumnWidthMap(columnWidth);
        sheet1.setHead(createTestListStringHead());
        //or 设置自适应宽度
        //sheet1.setAutoWidth(Boolean.TRUE);
        writer.write1(createTestListObject(), sheet1);

        //写第二个sheet sheet2  模型上打有表头的注解，合并单元格
        Sheet sheet2 = new Sheet(2, 3, WriteModel.class, "第二个sheet", null);
        sheet2.setTableStyle(createTableStyle());
        writer.write(createTestListJavaMode(), sheet2);

        //写第三个sheet包含多个table情况
        Sheet sheet3 = new Sheet(3, 0);
        sheet3.setSheetName("第三个sheet");
        Table table1 = new Table(1);
        table1.setHead(createTestListStringHead());
        writer.write1(createTestListObject(), sheet3, table1);

        //写sheet2  模型上打有表头的注解
        Table table2 = new Table(2);
        table2.setTableStyle(createTableStyle());
        table2.setClazz(WriteModel.class);
        writer.write(createTestListJavaMode(), sheet3, table2);

        writer.finish();
        out.close();
    }
}
