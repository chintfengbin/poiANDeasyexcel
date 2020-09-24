package com.fbin.easy;

import com.alibaba.excel.EasyExcel;
import org.junit.Test;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class TestEasy {

    String Path = "E:\\project\\poiANDeasyexcel\\poi\\";

    private List<DemoData> data(){
        List<DemoData> list = new ArrayList<DemoData>();
        for (int i= 0; i < 10 ; i++){
            DemoData data = new DemoData();
            data.setString("字符串" + i);
            data.setDate(new Date());
            data.setDoubleData(i+0.44);
            list.add(data);
        }
        return list;
    }

    @Test
    public void easyWrite(){
        //写法1
        String fileNmae = Path + "EasyTest.xlsx";
        //这里需要指定用哪个class去写，然后写道第一个sheet，名字为模板，然后文件流会自动关闭
        //如果这里想使用03 则 传入excelType参数即可
        EasyExcel.write(fileNmae,DemoData.class).sheet("模板1").doWrite(data());
    }

    @Test
    public void simpleRead() {
        // 有个很重要的点 DemoDataListener 不能被spring管理，要每次读取excel都要new,然后里面用到spring可以构造方法传进去
        // 写法1：
        String fileName = Path + "EasyTest.xlsx";
        // 这里 需要指定读用哪个class去读，然后读取第一个sheet 文件流会自动关闭
        EasyExcel.read(fileName, DemoData.class, new DemoDataListener()).sheet().doRead();

        /*// 写法2：
        fileName = TestFileUtil.getPath() + "demo" + File.separator + "demo.xlsx";
        ExcelReader excelReader = null;
        try {
            excelReader = EasyExcel.read(fileName, DemoData.class, new DemoDataListener()).build();
            ReadSheet readSheet = EasyExcel.readSheet(0).build();
            excelReader.read(readSheet);
        } finally {
            if (excelReader != null) {
                // 这里千万别忘记关闭，读的时候会创建临时文件，到时磁盘会崩的
                excelReader.finish();
            }
        }*/
    }
}
