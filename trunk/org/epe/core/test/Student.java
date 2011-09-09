package org.epe.core.test;

import org.epe.annotations.ECell;
import org.mec.validation.annotation.Phone;

public class Student {

	@ECell(name = "姓名")
	private String name;
	@ECell(name = "年龄")
	private Integer age;

	@ECell(name = "性别")
	private String sex;
	@ECell(name = "学号")
	private int sid;

	@Phone()
	private String phone;
	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public Integer getAge() {
		return age;
	}

	public void setAge(Integer age) {
		this.age = age;
	}

	public String getSex() {
		return sex;
	}

	public void setSex(String sex) {
		this.sex = sex;
	}

	public int getSid() {
		return sid;
	}

	public void setSid(int sid) {
		this.sid = sid;
	}

	public String getPhone() {
		return phone;
	}

	public void setPhone(String phone) {
		this.phone = phone;
	}

}
