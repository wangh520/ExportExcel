package com.twf.springcloud.ExportExcel.utils;

import com.twf.springcloud.ExportExcel.enumClass.ExcelFiled;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;

public class ExcelPOIUtils {
    final static String errormsg = "您好,第 [%s]行,第[%s]列%s,请重新输入!";
    private static Logger logger = Logger.getLogger(ExcelPOIUtils.class.getName());
    private final static String xls = "xls";
    private final static String xlsx = "xlsx";

    static SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

    SimpleDateFormat createDateFormat(String pattern) {
        return new SimpleDateFormat(pattern);
    }

    public static Workbook getWorkBook(MultipartFile file) throws Exception {
        // 获得文件名
        String fileName = file.getOriginalFilename();
        // 创建Workbook工作薄对象，表示整个excel
        Workbook workbook = null;
        try {
            // 获取excel文件的io流
            InputStream is = file.getInputStream();
            // 根据文件后缀名不同(xls和xlsx)获得不同的Workbook实现类对象
            if (fileName.endsWith(xls)) {
                // 2003
                workbook = new HSSFWorkbook(is);
            } else if (fileName.endsWith(xlsx)) {
                // 2007
                workbook = new XSSFWorkbook(is);
            }
        } catch (IOException e) {
            throw new Exception("解析报文出错!");
        }
        return workbook;
    }

    /**
     *
     * @param file   文件
     * @param class1    实体类
     * @param errorMap  错误map   key 行号   value  具体报错信息
     * @param errors    错误列表
     * @param colCount  列数
     * @return
     * @throws Exception
     */
    public static <T> List<T> readExc(MultipartFile file, Class<T> class1,Map<Integer, String> errorMap,List<T> errors,int colCount) throws Exception {
        try {
            Workbook workbook = getWorkBook(file);
            return readExcel(class1, errorMap,errors, workbook,colCount);
        } catch (Exception e) {
            throw new Exception(e.getMessage());
        }

    }

    public static <T> List<T> readExc(InputStream in, Class<T> class1,Map<Integer, String> errorMap,List<T> errors,int colCount) throws Exception {
        try {
            Workbook workbook = WorkbookFactory.create(in);
            return readExcel(class1, errorMap,errors, workbook,colCount);
        } catch (Exception e) {
            e.printStackTrace();
            throw new Exception(e.getMessage());
        }
    }

    private static <T> List<T> readExcel(Class<T> class1, Map<Integer, String> errorMap,List<T> errors, Workbook workbook,int colCount)
            throws Exception {
        Sheet sheet = workbook.getSheetAt(0);
        int rowCount = sheet.getPhysicalNumberOfRows(); // 获取总行数
        if (rowCount == 1) { // 表格最小长度应为2
            return null;
        }
        int coloumNum  =sheet.getRow(0).getPhysicalNumberOfCells();
        if(coloumNum != colCount){
            throw new Exception("excel格式跟模板格式不一致!");
        }


        List<T> list = new ArrayList<T>(rowCount - 1);
        T obj;
        // 遍历每一行
        for (int r = 1; r < rowCount; r++) {
            Row row = sheet.getRow(r);
            obj = class1.newInstance();
            Field[] fields = class1.getDeclaredFields();
            Field field;
            boolean flag = true;     //标识,确定该条数据是否通过第一轮判断
            boolean errorFlag = false;  //标识是否有错误点
            StringBuffer error = new StringBuffer();
            for (int j = 0; j < fields.length; j++) {
                field = fields[j];
                ExcelFiled excelFiled = field.getAnnotation(ExcelFiled.class);
                if(excelFiled == null){
                    continue;
                }
                Cell cell = row.getCell(excelFiled.colIndex());
                try {
                    //序号为0,编号为空,则为无效行
                    if(excelFiled.colIndex() == 0 ){
                        if(cell == null || cell.toString().length() == 0){
                            flag  = false;
                            break;
                        }
                    }

                    if(!excelFiled.skip()){   //必填字段,需要判断是否非空
                        if (cell == null || cell.toString().length() == 0) {
                            //errogMsg(errorMap, r, j, excelFiled,"不能为空");
                            error.append(excelFiled.colName()+"不能为空"+"|");
                            flag = false;
                            errorFlag = true;
                            continue;
                        }
                    }

                    if (field.getType().equals(Date.class)) {
                        if (CellType.STRING == cell.getCellType()) {
                            ReflectUtil.setValue(obj, field.getName(), dateFormat.parse(cell.getStringCellValue()));
                        } else {
                            ReflectUtil.setValue(obj, field.getName(), new Date(cell.getDateCellValue().getTime()));
                        }
                    } else if (field.getType().equals(java.lang.Integer.class)) {
                        if (CellType.NUMERIC == cell.getCellType()) {
                            ReflectUtil.setValue(obj, field.getName(), (int) cell.getNumericCellValue());
                        } else if (CellType.STRING == cell.getCellType()) {
                            ReflectUtil.setValue(obj, field.getName(), Integer.parseInt(cell.getStringCellValue()));
                        }
                    } else if (field.getType().equals(java.math.BigDecimal.class)) {
                        if (CellType.NUMERIC == cell.getCellType()) {
                            ReflectUtil.setValue(obj, field.getName(), new BigDecimal(cell.getNumericCellValue()));
                        } else if (CellType.STRING == cell.getCellType()) {
                            ReflectUtil.setValue(obj, field.getName(), cell.getStringCellValue());
                        }
                    }else if (field.getType().equals(java.lang.Double.class)) {
                        if (CellType.NUMERIC == cell.getCellType()) {
                            if(excelFiled.precision() == 0){   //没有小数点
                                ReflectUtil.setValue(obj, field.getName(),Double.valueOf(new BigDecimal(cell.getNumericCellValue()).intValue()));
                            }else{
                                ReflectUtil.setValue(obj, field.getName(),new BigDecimal(cell.getNumericCellValue()).doubleValue());
                            }
                        } else if (CellType.STRING == cell.getCellType()) {
                            ReflectUtil.setValue(obj, field.getName(), new Double(cell.getStringCellValue()));
                        }
                    } else if (field.getType().equals(java.lang.String.class)) {
                        if (CellType.NUMERIC == cell.getCellType()) {
                            if(excelFiled.precision() == 0){   //没有小数点
                                ReflectUtil.setValue(obj, field.getName(),new BigDecimal(cell.getNumericCellValue()).intValue()+"");
                            }else{
                                ReflectUtil.setValue(obj, field.getName(),new BigDecimal(cell.getNumericCellValue()).doubleValue()+"");
                            }
                        } else if (CellType.STRING == cell.getCellType()) {
                            ReflectUtil.setValue(obj, field.getName(), cell.getStringCellValue());
                        }
                    }  else {
                        if (CellType.NUMERIC == cell.getCellType()) {
                            ReflectUtil.setValue(obj, field.getName(), new BigDecimal(cell.getNumericCellValue()));
                        } else if (CellType.STRING == cell.getCellType()) {
                            ReflectUtil.setValue(obj, field.getName(), cell.getStringCellValue());
                        }
                    }
                } catch (Exception e) {
                    flag = false;
                    errorFlag = true;
                    error.append(excelFiled.colName()+"类型格式有误"+"|");
                }
            }
            //录入行号
//            ReflectUtil.setValue(obj, "rowNum",(r+1));
            if(flag){
                list.add(obj);
            }
            if(errorFlag){
                errorMap.put(r, error.toString());
                errors.add(obj);
            }
        }
        return list;
    }


}