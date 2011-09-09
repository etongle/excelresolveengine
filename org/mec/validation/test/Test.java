package org.mec.validation.test;

import java.util.Date;

import org.mec.validation.PatternCompile;

import com.itsv.gbp.core.util.DateFromat;

public class Test
{
	public static void main(String[] args)
	{

		Student student = new Student();
		student.setIdCard("4302251987121656110");

		student.setPhone("186108");

		student.setTel("0121-775825851");
		student.setPinteger(1234);
		student.setTestInte(12345677);

		PatternCompile pc = new PatternCompile();
		System.out.println(pc.compile(student, "idCard"));

		pc.compile(student, "phone");

		pc.compile(student, "tel");
		pc.compile(student, "tel");
		System.out.println(" ----" + pc.compile(student, "pinteger"));

		student.setTel("0121-77582585");
		student.setEmail("lishengbo4444@163.com");
		student.setPfloat(-28.0f);
		student.setPinteger(-100);
		student.setZipCode("2563145");
		student.setDate(DateFromat.formatDate("1987-12-16", "yyyy-MM-dd"));
		student.setDate(DateFromat.formatDate("1987-12asd-16", "yyyy-MM-dd"));

		System.out.println("身份证号:" + pc.compile(student, "idCard"));
		System.out.println("手机号码:" + pc.compile(student, "phone"));
		System.out.println("电话号码:" + pc.compile(student, "tel"));
		System.out.println("电子邮件:" + pc.compile(student, "email"));
		System.out.println("正整数:" + pc.compile(student, "pinteger"));
		System.out.println("正符点:" + pc.compile(student, "pfloat"));
		System.out.println("邮编:" + pc.compile(student, "zipCode"));
		System.out.println("日期验证" + pc.compile(student, "date"));

	}

}
