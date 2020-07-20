package com.zh;

import com.sun.jndi.ldap.Ber;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Date;

/**
 * @author Beloved
 * @date 2020/7/17 20:53
 * POI读Excel测试类
 */
public class ExcelReadTest {

    /**
     * 03简单的读操作
     */
    @Test
    public void ExcelRead03() throws IOException {


        String PATH = "E:\\ideaProject\\POIAndEasyExcel\\POI学生表.xls";

        //获取文件流
        FileInputStream inputStream = new FileInputStream(PATH);

        //创建工作簿  使用excel能操作这里都可以操作
        Workbook workbook = new HSSFWorkbook(inputStream);
        //根据下标获取工作表  也可以根据名
        Sheet sheet = workbook.getSheetAt(0);
        //获取行
        Row row = sheet.getRow(1);
        //获取列
        Cell cell = row.getCell(1);
        //获取值
        System.out.println(cell.getStringCellValue());
        inputStream.close();
    }

    /**
     * 07简单的读操作
     */
    @Test
    public void ExcelRead07() throws IOException {


        String PATH = "E:\\ideaProject\\POIAndEasyExcel\\POI学生表07.xlsx";

        //获取文件流
        FileInputStream inputStream = new FileInputStream(PATH);

        //创建工作簿  使用excel能操作这里都可以操作
        Workbook workbook = new XSSFWorkbook(inputStream);
        //根据下标获取工作表  也可以根据名
        Sheet sheet = workbook.getSheetAt(0);
        //获取行
        Row row = sheet.getRow(2);
        //获取列
        Cell cell = row.getCell(1);
        //获取值
        System.out.println(cell.getStringCellValue());
        inputStream.close();
    }

    /**
     * 测试读取不同的类型
     */
    @Test
    public void CellType() throws IOException {
        String PATH = "E:\\ideaProject\\POIAndEasyExcel\\明细表.xlsx";


        //获取文件流
        FileInputStream inputStream = new FileInputStream(PATH);

        //创建工作簿
        Workbook workbook = new XSSFWorkbook(inputStream);

        //获取表
        Sheet sheet = workbook.getSheetAt(0);

        // 获取标题内容
        // 获取第一行
        Row title = sheet.getRow(0);
        if (title != null){
            // 获取这一行有多少列
            int cellCount = title.getPhysicalNumberOfCells();
            for (int i = 0; i < cellCount; i++) {
                // 获取列
                Cell cell = title.getCell(i);
                if (cell!=null){
                    // 获取类型
                    int cellType = cell.getCellType();
                    // 获取值
                    String cellValue = cell.getStringCellValue();
                    System.out.print(cellValue + " | ");
                }
            }
            System.out.println();
        }

        // 获取表中的内容
        // 获取所有行
        int rowCount = sheet.getPhysicalNumberOfRows();
        for (int rowNum = 1; rowNum < rowCount; rowNum++) {
            // 获取行
            Row row = sheet.getRow(rowNum);
            if (row!=null){
                // 获取这一行有多少列
                int cellCount = title.getPhysicalNumberOfCells();
                for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                    System.out.print("[" + (rowNum+1) + "-" + (cellNum+1) + "]");
                    // 读取列
                    Cell cell = row.getCell(cellNum);
                    // 匹配列的数据类型
                    if (cell != null){
                        int cellType = cell.getCellType();

                        switch (cellType){
                            case XSSFCell.CELL_TYPE_STRING: // 字符串
                                System.out.print(" 字符串 ");
                                System.out.println(cell.getStringCellValue());
                                break;
                            case XSSFCell.CELL_TYPE_BOOLEAN: // 布尔类型
                                System.out.print(" 布尔类型 ");
                                System.out.println(cell.getBooleanCellValue());
                                break;
                            case XSSFCell.CELL_TYPE_BLANK: // 空值
                                System.out.println(" 空值 ");
                                break;
                            case XSSFCell.CELL_TYPE_NUMERIC: // 数字（日期、普通数字）
                                if (HSSFDateUtil.isCellDateFormatted(cell)){  // 日期
                                    Date date = cell.getDateCellValue();
                                    System.out.print(" 日期 ");
                                    System.out.println(new DateTime(date).toString("yyyy-MM-dd HH:mm:ss"));
                                }else{
                                    // 如果不是日期格式，防止数字过长,设置为字符串 可以读取科学计数法
                                    cell.setCellType(XSSFCell.CELL_TYPE_STRING);
                                    System.out.print(" 数字 ");
                                    System.out.println(cell.getStringCellValue());
                                }
                                break;
                            case XSSFCell.CELL_TYPE_ERROR: // 错误
                                System.out.println("数据类型错误 ");
                                break;
                        }
                    }
                }
            }
        }

        inputStream.close();
    }

    /**
     * 读取计算公式
     */
    @Test
    public void Formula() throws IOException {
        String PATH = "E:\\ideaProject\\POIAndEasyExcel\\公式.xlsx";

        //获取文件流
        FileInputStream inputStream = new FileInputStream(PATH);

        //创建工作簿
        Workbook workbook = new XSSFWorkbook(inputStream);

        //获取表
        Sheet sheet = workbook.getSheetAt(0);

        Row row = sheet.getRow(3);
        Cell cell = row.getCell(0);


        // 拿到计算公式 eval
        FormulaEvaluator formulaEvaluator = new XSSFFormulaEvaluator((XSSFWorkbook)workbook);

        //输出单元格内容
        int cellType = cell.getCellType();
        switch (cellType){
            case Cell.CELL_TYPE_FORMULA: //公式
                // 获取公式
                String formula = cell.getCellFormula();
                System.out.println(formula);

                // 计算结果
                CellValue cellValue = formulaEvaluator.evaluate(cell);
                String value = cellValue.formatAsString();
                System.out.println(value);
                break;
        }
        inputStream.close();
    }
}
