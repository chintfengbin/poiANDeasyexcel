package com.fbin;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileInputStream;
import java.util.Date;


public class ExcelReadTest {
    public String Path = "E:\\project\\poiANDeasyexcel\\poi";

    @Test
    public void testRead03() throws Exception{

        //获取文件流
        FileInputStream fileInputStream = new FileInputStream(Path+ "冯斌信息表03.xls");

        //创建一个工作簿，时excel中能操作的这边都能操作！
        Workbook workbook = new HSSFWorkbook(fileInputStream);

        //获取信息:得到表
        Sheet sheet = workbook.getSheetAt(0);
        //得到行
        Row row = sheet.getRow(0);
        //得到列
        Cell cell = row.getCell(1);
        System.out.println(cell.getStringCellValue());
        fileInputStream.close();
    }

    @Test
    public void testRead07() throws Exception{

        //获取文件流
        FileInputStream fileInputStream = new FileInputStream(Path+ "冯斌信息表07.xlsx");

        //创建一个工作簿，时excel中能操作的这边都能操作！
        Workbook workbook = new XSSFWorkbook(fileInputStream);

        //获取信息:得到表
        Sheet sheet = workbook.getSheetAt(0);
        //得到行
        Row row = sheet.getRow(0);
        //得到列
        Cell cell = row.getCell(1);
        System.out.println(cell.getStringCellValue());
        fileInputStream.close();
    }

    @Test
    public void testReadDifcult07() throws Exception{

        String Path = "E:\\project\\poiANDeasyexcel\\";
        //获取文件流
        FileInputStream fileInputStream = new FileInputStream(Path+ "多类型测试用表07.xlsx");

        //创建一个工作簿，时excel中能操作的这边都能操作！
        Workbook workbook = new XSSFWorkbook(fileInputStream);

        //获取标题内容
        Sheet sheet = workbook.getSheetAt(0);

        //获取标题内容
        Row title = sheet.getRow(0);
        if (title != null){
            int cellCount = title.getPhysicalNumberOfCells();
            for (int cellNum = 0;cellNum<cellCount;cellNum++){
                Cell cell= title.getCell(cellNum);
                if (cell != null){
                    int cellType = cell.getCellType();
                    String cellValue = cell.getStringCellValue();
                    System.out.print(cellValue + " | ");
                }
            }
            System.out.println();
        }

        //获取计算公司eval
        FormulaEvaluator formulaEvaluator = new XSSFFormulaEvaluator((XSSFWorkbook) workbook);
        //获取表中内容
        //获取行数
        int roeCount = sheet.getPhysicalNumberOfRows();

        for (int rowNum =1; rowNum<roeCount;rowNum++){
            Row row = sheet.getRow(rowNum);
            if (row != null){

                for (int cellNum =0;cellNum<title.getPhysicalNumberOfCells();cellNum++){

                    System.out.print("["+(rowNum+1)+"-"+(cellNum+1)+"]");
                    Cell cell = row.getCell(cellNum);

                    //匹配数据类型
                    if (cell != null){

                        int cellType = cell.getCellType();
                        Object cellValue = "";

                        switch (cellType){

                            case Cell.CELL_TYPE_STRING://字符串
                                System.out.print("【String】");
                                cellValue = cell.getStringCellValue();
                                break;

                            case Cell.CELL_TYPE_NUMERIC://数字（日期等）
                                System.out.print("【NUMERIC】");
                                if (HSSFDateUtil.isCellDateFormatted(cell)){//日期,普通数字
                                    System.out.print("【日期】");
                                    Date date = cell.getDateCellValue();
                                    cellValue = new DateTime(date).toString("yyyy-MM-dd");
                                }else {
                                    System.out.print("【不是日期转换为字符串】");
                                    cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                                    cellValue= cell.toString();
                                }
                                break;

                            case Cell.CELL_TYPE_BLANK://空
                                System.out.print("【BLANK】");
                                break;

                            case Cell.CELL_TYPE_BOOLEAN://布尔
                                System.out.print("【布尔】");
                                cellValue = String.valueOf(cell.getBooleanCellValue());
                                break;

                            case Cell.CELL_TYPE_ERROR:
                                System.out.print("【错误类型】");
                                break;

                            case Cell.CELL_TYPE_FORMULA:
                                CellValue evaluate = formulaEvaluator.evaluate(cell);
                                cellValue = (int)(evaluate.getNumberValue()+0.5);//四舍五入

                        }

                        System.out.println(cellValue);
                    }

                }
            }

        }

        fileInputStream.close();
    }
}
