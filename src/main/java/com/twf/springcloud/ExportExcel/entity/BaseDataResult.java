package com.twf.springcloud.ExportExcel.entity;

/**
 * @author wangh
 * @2020/8/12 - 17:11
 */
public class BaseDataResult {

    private String code;
    private String mesage;

    public BaseDataResult(String s, String message) {

    }

    public String getCode() {
        return code;
    }

    public void setCode(String code) {
        this.code = code;
    }

    public String getMesage() {
        return mesage;
    }

    public void setMesage(String mesage) {
        this.mesage = mesage;
    }
}
