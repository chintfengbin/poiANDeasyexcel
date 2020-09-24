package com.fbin;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
//Hssf比较快，但是数据少
//Xssf比较慢但是数据多
//Sxssf兼具快和多，但是会产生缓存要在结束时清理
public class ExcelWriterTest {

    public String Path = "E:\\project\\poiANDeasyexcel\\poi";

    @Test
    public void writeTest03() throws Exception {
        //创建工作簿03
        Workbook workbook = new HSSFWorkbook();
        //创建sheet表
        Sheet sheet = workbook.createSheet("冯斌测试用表");
        //创建行,0代表第一行
        Row row1 = sheet.createRow(0);
        //创建列，第一行第一列
        Cell cell11 = row1.createCell(0);
        //给第一行第一列赋值
        cell11.setCellValue("冯斌学历");

        Cell cell12 = row1.createCell(1);
        cell12.setCellValue("本科");

        //创建行,1代表第二行
        Row row2 = sheet.createRow(1);
        //创建列，第二行第一列
        Cell cell21 = row2.createCell(0);
        //给第一行第一列赋值
        cell21.setCellValue("冯斌身高");

        Cell cell22 = row2.createCell(1);
        cell22.setCellValue("171");

        //创建一张表（IO流）
        FileOutputStream fileOutputStream = new FileOutputStream(Path + "冯斌信息表03.xls");

        //写入磁盘
        workbook.write(fileOutputStream);

        //关闭流
        fileOutputStream.close();

    }

    @Test
    public void writeTest07() throws Exception {
        //创建工作簿07
        Workbook workbook = new XSSFWorkbook();
        //创建sheet表
        Sheet sheet = workbook.createSheet("冯斌测试用表");
        //创建行,0代表第一行
        Row row1 = sheet.createRow(0);
        //创建列，第一行第一列
        Cell cell11 = row1.createCell(0);
        //给第一行第一列赋值
        cell11.setCellValue("冯斌学历");

        Cell cell12 = row1.createCell(1);
        cell12.setCellValue("本科");

        //创建行,1代表第二行
        Row row2 = sheet.createRow(1);
        //创建列，第二行第一列
        Cell cell21 = row2.createCell(0);
        //给第一行第一列赋值
        cell21.setCellValue("冯斌身高");

        Cell cell22 = row2.createCell(1);
        cell22.setCellValue("171");

        //创建一张表（IO流）
        FileOutputStream fileOutputStream = new FileOutputStream(Path + "冯斌信息表07.xlsx");

        //写入磁盘
        workbook.write(fileOutputStream);

        //关闭流
        fileOutputStream.close();

    }

    @Test
    public void bigDataWriteTest03() throws Exception{
        //时间
        long begin = System.currentTimeMillis();

        //创建一个工作簿03
        Workbook workbook = new HSSFWorkbook();
        //创建表
        Sheet sheet = workbook.createSheet("大数据测试用表");
        //插入大数据
        for (int row =0;row<65536;row++){
            Row row1 = sheet.createRow(row);
            for (int cell = 0;cell<=10;cell++){
                Cell cell1 = row1.createCell(cell);
                cell1.setCellValue(cell);
            }
        }
        System.out.println("over");
        FileOutputStream fileOutputStream = new FileOutputStream(Path + "大数据测试03.xls");

        workbook.write(fileOutputStream);

        long endTime = System.currentTimeMillis();
        double usingTime = (double)(endTime-begin)/1000;
        System.out.println(usingTime);

        fileOutputStream.close();
    }

    @Test
    public void bigDataWriteTest07() throws Exception{
        //时间
        long begin = System.currentTimeMillis();

        //创建一个工作簿07
        Workbook workbook = new XSSFWorkbook();
        //创建表
        Sheet sheet = workbook.createSheet("大数据测试用表");
        //插入大数据
        for (int row =0;row<65539;row++){
            Row row1 = sheet.createRow(row);
            for (int cell = 0;cell<=10;cell++){
                Cell cell1 = row1.createCell(cell);
                cell1.setCellValue(cell);
            }
        }
        System.out.println("over");
        FileOutputStream fileOutputStream = new FileOutputStream(Path + "大数据测试07.xlsx");

        workbook.write(fileOutputStream);

        long endTime = System.currentTimeMillis();
        double usingTime = (double)(endTime-begin)/1000;
        System.out.println(usingTime);

        fileOutputStream.close();
    }

    @Test
    public void bigDataWriteTestS() throws Exception{
        //时间
        long begin = System.currentTimeMillis();

        //创建一个工作簿S
        Workbook workbook = new SXSSFWorkbook();
        //创建表
        Sheet sheet = workbook.createSheet("大数据测试用表");
        //插入大数据
        for (int row =0;row<65539;row++){
            Row row1 = sheet.createRow(row);
            for (int cell = 0;cell<=10;cell++){
                Cell cell1 = row1.createCell(cell);
                cell1.setCellValue(cell);
            }
        }
        System.out.println("over");
        FileOutputStream fileOutputStream = new FileOutputStream(Path + "大数据测试s.xlsx");

        workbook.write(fileOutputStream);
        //清除临时文件（缓存）
        ((SXSSFWorkbook) workbook).dispose();
        long endTime = System.currentTimeMillis();
        double usingTime = (double)(endTime-begin)/1000;
        System.out.println(usingTime);

        fileOutputStream.close();
    }
}
