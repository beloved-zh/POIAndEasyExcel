package com.zh.pojo;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.format.DateTimeFormat;
import com.zh.converter.SexConverter;
import lombok.Data;

import java.util.Date;

/**
 * @author Beloved
 * @date 2020/7/17 22:40
 */
@Data
public class User {

    // 指定表头 可以通过列名或这下标 但是不能一起使用
    // @ExcelProperty(index = 1)
    @ExcelProperty("学号")
    private String id;
    @ExcelProperty("姓名")
    private String name;
    @ExcelProperty("年龄")
    private Integer age;
    @ExcelProperty(value = "性别",converter = SexConverter.class)
    private Integer sex;
    @ExcelProperty("班级")
    private String clas;
    @ExcelProperty("电话")
    private String tel;
    // 这个注解必须是String类型
    @DateTimeFormat("yyyy-MM-dd HH:mm:ss")
    @ExcelProperty("生日")
    private String date;

}
