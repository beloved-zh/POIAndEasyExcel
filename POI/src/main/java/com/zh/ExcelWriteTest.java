package com.zh;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @author Beloved
 * @date 2020/7/17 19:10
 * POI写Excel测试类
 */
public class ExcelWriteTest {

    String PATH = "E:\\ideaProject\\POIAndEasyExcel\\POI";

    /**
     * 03版本写入测试
     */
    @Test
    public void ExcelWrite03() throws IOException {
        //1.创建一个工作簿
        Workbook workbook = new HSSFWorkbook();
        //2.创建一个工作表
        Sheet sheet = workbook.createSheet("学生表");
        //3.创建一个行(第一行)
        Row row1 = sheet.createRow(0);
        //4.创建单元格
        //(1,1)
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("姓名");
        //(1,2)
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue("张恒");

        //第二行
        Row row2 = sheet.createRow(1);
        //(2,1)
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("年龄");
        //(2,2)
        Cell cell22 = row2.createCell(1);
        cell22.setCellValue(20);

        //第三行
        Row row3 = sheet.createRow(2);
        //(3,1)
        Cell cell31 = row3.createCell(0);
        cell31.setCellValue("生日");
        //(3,2)
        Cell cell32 = row3.createCell(1);
        cell32.setCellValue(new DateTime().toString("yyyy-MM-dd HH:mm:ss"));

        //5.生成表（IO流）    03版本使用.xls结尾
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "学生表03.xls");

        //6.输出
        workbook.write(fileOutputStream);

        //7.关闭流
        fileOutputStream.close();

        System.out.println("生成完毕");
    }


    /**
     * 07版本写入测试
     */
    @Test
    public void ExcelWrite07() throws IOException {
        //1.创建一个工作簿
        Workbook workbook = new XSSFWorkbook();
        //2.创建一个工作表
        Sheet sheet = workbook.createSheet("学生表");
        //3.创建一个行(第一行)
        Row row1 = sheet.createRow(0);
        //4.创建单元格
        //(1,1)
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("姓名");
        //(1,2)
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue("张恒");

        //第二行
        Row row2 = sheet.createRow(1);
        //(2,1)
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("年龄");
        //(2,2)
        Cell cell22 = row2.createCell(1);
        cell22.setCellValue(20);

        //第三行
        Row row3 = sheet.createRow(2);
        //(3,1)
        Cell cell31 = row3.createCell(0);
        cell31.setCellValue("生日");
        //(3,2)
        Cell cell32 = row3.createCell(1);
        cell32.setCellValue(new DateTime().toString("yyyy-MM-dd HH:mm:ss"));

        //5.生成表（IO流）    07版本使用.xlsx结尾
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "学生表07.xlsx");

        //6.输出
        workbook.write(fileOutputStream);

        //7.关闭流
        fileOutputStream.close();

        System.out.println("生成完毕");
    }


    /**
     * 03版本大数据写入测试
     */
    @Test
    public void ExcelWrite03BigData() throws IOException {
        //开始时间
        long begin = System.currentTimeMillis();

        //创建工作簿
        Workbook workbook = new HSSFWorkbook();
        //创建表
        Sheet sheet = workbook.createSheet();
        //写入数据
        for (int i = 0; i < 65537; i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < 10; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(j);
            }
        }

        System.out.println("over");

        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "ExcelWrite03BigData.xls");

        workbook.write(fileOutputStream);

        fileOutputStream.close();

        //结束时间
        long end = System.currentTimeMillis();

        System.out.println((double) (end - begin) / 1000);
    }

    /**
     * 07版本大数据写入测试
     */
    @Test
    public void ExcelWrite07BigData() throws IOException {
        //开始时间
        long begin = System.currentTimeMillis();

        //创建工作簿
        Workbook workbook = new XSSFWorkbook();
        //创建表
        Sheet sheet = workbook.createSheet();
        //写入数据
        for (int i = 0; i < 100000; i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < 10; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(j);
            }
        }

        System.out.println("over");

        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "ExcelWrite07BigData.xlsx");

        workbook.write(fileOutputStream);

        fileOutputStream.close();

        //结束时间
        long end = System.currentTimeMillis();

        System.out.println((double) (end - begin) / 1000);
    }


    /**
     * SXSSF大数据写入测试
     */
    @Test
    public void ExcelWrite07BigDataS() throws IOException {
        //开始时间
        long begin = System.currentTimeMillis();

        //创建工作簿
        Workbook workbook = new SXSSFWorkbook();
        //创建表
        Sheet sheet = workbook.createSheet();
        //写入数据
        for (int i = 0; i < 100000; i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < 10; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(j);
            }
        }

        System.out.println("over");

        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "ExcelWrite07BigDataS.xlsx");

        workbook.write(fileOutputStream);

        fileOutputStream.close();

        //清除临时文件
        ((SXSSFWorkbook)workbook).dispose();

        //结束时间
        long end = System.currentTimeMillis();

        System.out.println((double) (end - begin) / 1000);
    }
}
