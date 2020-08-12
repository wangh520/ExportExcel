package com.twf.springcloud.ExportExcel.controller;


import com.twf.springcloud.ExportExcel.po.User;
import com.twf.springcloud.ExportExcel.utils.ExcelPOIUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

public class Main {
    public static void main(String[] args) {
        try {
            InputStream in = new FileInputStream(new File("D://1.xlsx"));
            //错误列表
            HashMap<Integer, String> errorMap = new HashMap<>();
            List<User> errors = new ArrayList<User>();
            List<User> readExc = ExcelPOIUtils.readExc(in, User.class,errorMap,errors,6);
            //List<ExcelEntity> readExc = ExcelPOIUtils.readExc(in, ExcelEntity.class,errorMap,3);

            System.out.println("数据----------------");
            for (User s : readExc) {
                System.out.println(s.toString());
            }

            System.out.println("错误maps----------------");
            errorMap.forEach((k,v)-> System.out.println(v));

            System.out.println("错误列表----------------");
            errors.forEach(k-> System.out.println(k.toString()));

            int success = readExc.size();
            int error = errorMap.size();
            System.out.println("总数"+(success+error)+",成功数:"+success+",失败数:"+errorMap.size());
        }  catch (Exception e) {
            e.printStackTrace();
            System.out.println(e.getMessage());
        }
    }
}