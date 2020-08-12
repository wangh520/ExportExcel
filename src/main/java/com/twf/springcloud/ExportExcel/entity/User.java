package com.twf.springcloud.ExportExcel.entity;

import com.twf.springcloud.ExportExcel.enumClass.ExcelFiled;

import java.io.Serializable;
/**
 * @Author wangh
 * @Description //TODO
 * @Date 13:57 2020/8/12
 * @Param 
 * @return 
 **/
public class User implements Serializable{
	@ExcelFiled(colIndex=0,colName="姓名",skip=true)
	private String name; // 姓名
	@ExcelFiled(colIndex=1,colName="性别")
	private String sex; // 性别
	@ExcelFiled(colIndex=2,colName="年龄")
	private Integer age; // 年龄
	@ExcelFiled(colIndex=3,colName="手机号",skip=true)
	private String phoneNo; // 手机号
	@ExcelFiled(colIndex=4,colName="地址",skip=true)
	private String address; // 地址
	@ExcelFiled(colIndex=5,colName="爱好")
	private String hobby; // 爱好
	
	public User(String name, String sex, Integer age, String phoneNo, String address, String hobby) {
		super();
		this.name = name;
		this.sex = sex;
		this.age = age;
		this.phoneNo = phoneNo;
		this.address = address;
		this.hobby = hobby;
	}

	public User() {

	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getSex() {
		return sex;
	}

	public void setSex(String sex) {
		this.sex = sex;
	}

	public Integer getAge() {
		return age;
	}

	public void setAge(Integer age) {
		this.age = age;
	}

	public String getPhoneNo() {
		return phoneNo;
	}

	public void setPhoneNo(String phoneNo) {
		this.phoneNo = phoneNo;
	}

	public String getAddress() {
		return address;
	}

	public void setAddress(String address) {
		this.address = address;
	}

	public String getHobby() {
		return hobby;
	}

	public void setHobby(String hobby) {
		this.hobby = hobby;
	}

	@Override
	public String toString() {
		return "User [name=" + name + ", sex=" + sex + ", age=" + age + ", phoneNo=" + phoneNo + ", address=" + address
				+ ", hobby=" + hobby + "]";
	}
}
