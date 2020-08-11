package com.twf.springcloud.ExportExcel.po;

import java.io.Serializable;

public class User implements Serializable{

	private static final long serialVersionUID = -9180229310895087286L;

	private String name; // 姓名
	
	private String sex; // 性别
	
	private Integer age; // 年龄
	
	private String phoneNo; // 手机号
	
	private String address; // 地址
	
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
