package com.hhawking.parctical.tools;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
* @title: 读取Excel转成POJO
* @Author: HH
* @Date: 2019/6/4 11:04
*/
public class ExcelUtils{

    public static <T> List<T> readExcelToPojo(File excel, Map<String,String> map, Class<T> tClass) throws Exception{
        //判断文件是否存在
        if (excel.isFile() && excel.exists()) {
            //.是特殊字符，需要转义
            String[] split = excel.getName().split("\\.");
            Workbook wb;
            //根据文件后缀（xls/xlsx）进行判断
            if ( "xls".equals(split[split.length-1])){
                FileInputStream fileInputStream = new FileInputStream(excel);   //文件流对象
                wb = new HSSFWorkbook(fileInputStream);
            }else if ("xlsx".equals(split[split.length-1])){
                wb = new XSSFWorkbook(excel);
            }else {
                return null;
            }
            return read(wb,map,tClass);
        }
        return null;
    }

    public static <T> List<T> excelToPojo(Workbook wb, Map<String,String> map, Class<T> tClass) throws Exception{
        return read(wb, map, tClass);
    }

    private static <T> List<T> read(Workbook wb, Map<String, String> map, Class<T> tClass) throws InstantiationException, IllegalAccessException, NoSuchFieldException {
        List<T> list = new ArrayList<>();

        //开始解析,读取sheet 0
        Sheet sheet = wb.getSheetAt(0);
        Row titles = sheet.getRow(1);
        int cells = titles.getPhysicalNumberOfCells();
        //去掉空白列
        for (int i = 0; i < cells; i++) {
            String s = titles.getCell(i).toString();
            if (StringUtils.isBlank(s)){
                cells = i;
            }
        }

        //前2行是标题和列名，所以不读
        int firstRowIndex = sheet.getFirstRowNum()+2;
        int lastRowIndex = sheet.getLastRowNum();

        //遍历行
        for(int rIndex = firstRowIndex; rIndex <= lastRowIndex; rIndex++) {
            Row row = sheet.getRow(rIndex);
            if (row != null) {
                T voWorkUnit = tClass.newInstance();
                for (int i = 0; i < cells; i++) {
                    Cell cell = row.getCell(i);
                    if (cell != null) {
                        //由于读取手机号码时，POI会将其转为科学计数法，导致数据出错，所以这里要设置格式
                        cell.setCellType(CellType.STRING);
                        String s = cell.toString();
                        String property = map.get(titles.getCell(i).toString());
                        if (!StringUtils.isBlank(property)) {
                            Field field = tClass.getDeclaredField(property);
                            field.setAccessible(true);
                            field.set(voWorkUnit,s);

                        }
                    }
                }
                list.add(voWorkUnit);
            }
        }

        return list;
    }

    public static <T> HSSFWorkbook pojoToExcel(List<T> list, Map<String, String> map) {
        Set<String> set = map.keySet();
        List<String> keyList = new ArrayList<>(set);
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("用户信息");
        HSSFRow title = sheet.createRow(0);

        for (int i = 0; i < keyList.size(); i++) {
            title.createCell(i).setCellValue(keyList.get(i));
        }

        for (int i = 0; i < list.size(); i++) {
            T t = list.get(i);
            HSSFRow row = sheet.createRow(i+1);
            for (int j = 0; j < keyList.size(); j++) {
                String key = keyList.get(j);
                String property = map.get(key);
                // 通过属性获取对象的属性
                try {
                    Field field = t.getClass().getDeclaredField(property);
                    // 对象的属性的访问权限设置为可访问
                    field.setAccessible(true);
                    // 获取属性的对应的值
                    Object o = field.get(t);
                    String value = "";
                    if (o!=null)
                        value = o.toString();
                    //写入单元格
                    row.createCell(j).setCellValue(value);
                } catch (NoSuchFieldException | IllegalAccessException e) {
                    e.printStackTrace();
                }
            }
        }
        return workbook;
    }
}
