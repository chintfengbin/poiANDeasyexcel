package com.fbin.easy;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

import java.util.Date;

@Data
public class DemoData {

    @ExcelProperty("字符串")
    private String string;

    @ExcelProperty("日期")
    private Date date;

    @ExcelProperty("数字")
    private Double doubleData;


}
