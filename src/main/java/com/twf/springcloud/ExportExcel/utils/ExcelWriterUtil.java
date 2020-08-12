package com.twf.springcloud.ExportExcel.utils;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.util.IOUtils;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.util.StringUtils;

import java.io.InputStream;
import java.net.URLEncoder;
import java.util.HashMap;
import java.util.Map;

/**
 * @Author wangh
 * @Description //TODO 
 * @Date 12:01 2020/8/12
 * @Param 
 * @return 
 **/
public class ExcelWriterUtil {


    /**
     * 生成sheet表，并写入第一行数据（列头）
     * @param workbook 工作簿对象
     * @return 已经写入列头的Sheet
     */
    public static Sheet buildDataSheet(Workbook workbook,String[] title,int[] titleLength,String tableName) {
        Sheet sheet = workbook.createSheet();
        // 设置默认行高
        sheet.setDefaultRowHeight((short) 400);
        // 构建头单元格样式
        CellStyle cellStyle = buildHeadCellStyle(sheet.getWorkbook());

        // 合表名并行(4个参数，分别为起始行，结束行，起始列，结束列)
        CellRangeAddress region = new CellRangeAddress(0, 0, 0, title.length-1);
        sheet.addMergedRegion(region);
        //第一行添加表名
        Row row0 = sheet.createRow(0);
        row0.setHeight((short) 500);
        Cell cell0 = row0.createCell(0);
        cell0.setCellValue(tableName);
        cell0.setCellStyle(cellStyle);
        setRegionBorderStyle(region, sheet);   //给合并过的单元格加边框

        // 写入第二行各列的数据
        Row head = sheet.createRow(1);
        for (int i = 0; i < title.length; i++) {
            Cell cell = head.createCell(i);
            cell.setCellValue(title[i]);
            cell.setCellStyle(cellStyle);
            sheet.setColumnWidth(i, titleLength[i]);// 设置列头宽度
        }
        return sheet;
    }

    /**
     * 设置第一行列头的样式
     * @param workbook 工作簿对象
     * @return 单元格样式对象
     */
    private static CellStyle buildHeadCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        //对齐方式设置
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);// 设置单元格字体显示居中(上下方向)
        // 设置自动换行
        style.setWrapText(true);
        //边框颜色和宽度设置
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); // 下边框
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); // 左边框
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); // 右边框
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex()); // 上边框
        //设置背景颜色
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //粗体字设置
        Font font = workbook.createFont();
        font.setBold(true);
//        font.setFontHeightInPoints((short) 10);// 设置字体的高度
        style.setFont(font);
        return style;
    }

    /**
     * 设置报表体样式
     * @param workbook
     * @return
     */
    public static CellStyle contentStyle(Workbook workbook){
        // 设置style1的样式，此样式运用在第二行
        CellStyle style1 = workbook.createCellStyle();// cell样式
        // 设置单元格上、下、左、右的边框线
        style1.setBorderBottom(BorderStyle.THIN);
        style1.setBorderLeft(BorderStyle.THIN);
        style1.setBorderRight(BorderStyle.THIN);
        style1.setBorderTop(BorderStyle.THIN);
        style1.setWrapText(false);// 设置自动换行
        style1.setAlignment(HorizontalAlignment.CENTER);// 设置单元格字体显示居中（左右方向）
        style1.setVerticalAlignment(VerticalAlignment.CENTER);// 设置单元格字体显示居中(上下方向)
        return style1;
    }


    /**
     * 下载文件,纯SpringMVC的API来完成
     *
     * @param is
     *            文件输入流
     * @param name
     *            文件名称,带后缀名
     *
     * @throws Exception
     */
    public static ResponseEntity<byte[]> buildResponseEntity(InputStream is, String name) throws Exception {
        HttpHeaders header = new HttpHeaders();
        String fileSuffix = name.substring(name.lastIndexOf('.') + 1);
        fileSuffix = fileSuffix.toLowerCase();

        Map<String, String> arguments = new HashMap<String, String>();
        arguments.put("xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        arguments.put("xls", "application/vnd.ms-excel");

        String contentType = arguments.get(fileSuffix);
        header.add("Content-Type", (StringUtils.hasText(contentType) ? contentType : "application/x-download"));
        if(is!=null && is.available()!=0){
            header.add("Content-Length", String.valueOf(is.available()));
            header.add("Content-Disposition", "attachment;filename*=utf-8'zh_cn'" + URLEncoder.encode(name, "UTF-8"));
            byte[] bs = IOUtils.toByteArray(is);
            return new ResponseEntity(bs, header, HttpStatus.OK);
        }else{
            String string="数据为空";
            header.add("Content-Length", "0");
            header.add("Content-Disposition", "attachment;filename*=utf-8'zh_cn'" + URLEncoder.encode(name, "UTF-8"));
            return new ResponseEntity(string.getBytes(), header, HttpStatus.OK);
        }
    }

    /**
     * @Author wangh
     * @Description //TODO 给合并后的单元格 加边框
     * @Date 10:02 2020/8/12
     * @Param [region, sheet]
     * @return void
     **/
    public static void setRegionBorderStyle(CellRangeAddress region, Sheet sheet){
        RegionUtil.setBorderBottom(BorderStyle.THIN, region, sheet);//下边框
        RegionUtil.setBorderLeft(BorderStyle.THIN, region, sheet);     //左边框
        RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet);    //右边框
        RegionUtil.setBorderTop(BorderStyle.THIN, region, sheet);      //上边框
    }
}
