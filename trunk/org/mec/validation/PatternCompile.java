package org.mec.validation;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.log4j.Logger;
import org.epe.core.test.Student;
import org.mec.validation.annotation.Date;
import org.mec.validation.annotation.Email;
import org.mec.validation.annotation.IDCard;
import org.mec.validation.annotation.PFloat;
import org.mec.validation.annotation.PInteger;
import org.mec.validation.annotation.Phone;
import org.mec.validation.annotation.Tel;
import org.mec.validation.annotation.ZipCode;

import com.itsv.gbp.core.util.DateFromat;

/**
 * ��֤��,
 * 
 * @author ������
 * 
 */
public class PatternCompile
{

	private Logger log = Logger.getLogger(this.getClass());

	public boolean compile(String str,String reg)
	{
		Pattern pattern=Pattern.compile(reg);
		Matcher matcher=pattern.matcher(str);
		return matcher.matches();
	}
	/**
	 * ������ʽ��֤��
	 * 
	 * @param t
	 *            ʵ������
	 * @param field
	 *            �ֶ�
	 * @return
	 */
	public <T> boolean compile(T t, String field)
	{
		boolean b = false;
		try
		{
			Field field_new = null;
			field_new = t.getClass().getDeclaredField("" + field + "");

			field_new.setAccessible(true);
			Annotation[] anno = field_new.getDeclaredAnnotations();
			for (Annotation annotation : anno)
			{
				if (annotation instanceof IDCard)
				{
					Object obj = field_new.get(t);
					Pattern pattern = null;
					if (obj != null && obj.toString().length() == 18)
						pattern = java.util.regex.Pattern.compile(((IDCard) annotation).new_pattern());
					else
						pattern = java.util.regex.Pattern.compile(((IDCard) annotation).pattern());
					Matcher matcher = pattern.matcher(obj == null ? "" : obj.toString());
					b = matcher.matches();
					log.debug("�������֤��:" + b);
				}
				if (annotation instanceof Phone)
				{
					Object obj = field_new.get(t);
					Pattern pattern = null;
					pattern = java.util.regex.Pattern.compile(((Phone) annotation).pattern());
					Matcher matcher = pattern.matcher(obj == null ? "" : obj.toString());
					b = matcher.matches();
					log.debug("�����ֻ�����:" + b);
				}
				if (annotation instanceof Tel)
				{

					Object obj = field_new.get(t);
					Pattern pattern = null;
					pattern = java.util.regex.Pattern.compile(((Tel) annotation).pattern());
					Matcher matcher = pattern.matcher(obj == null ? "" : obj.toString());
					b = matcher.matches();
					log.debug("����绰����:" + b);
				}
				if (annotation instanceof Email)
				{

					Object obj = field_new.get(t);
					Pattern pattern = null;
					pattern = java.util.regex.Pattern.compile(((Email) annotation).pattern());
					Matcher matcher = pattern.matcher(obj == null ? "" : obj.toString());
					b = matcher.matches();
					log.debug("��������ʼ�:" + b);
				}
				if (annotation instanceof PInteger)
				{

					Object obj = field_new.get(t);
					Pattern pattern = null;
					pattern = java.util.regex.Pattern.compile(((PInteger) annotation).pattern());
					Matcher matcher = pattern.matcher(obj == null ? "" : obj.toString());
					b = matcher.matches();
					log.debug("����������:" + b);
				}
				if (annotation instanceof PFloat)
				{

					Object obj = field_new.get(t);
					Pattern pattern = null;
					pattern = java.util.regex.Pattern.compile(((PFloat) annotation).pattern());
					Matcher matcher = pattern.matcher(obj == null ? "" : obj.toString());
					b = matcher.matches();
					log.debug("������������:" + b);
				}
				if (annotation instanceof ZipCode)
				{

					Object obj = field_new.get(t);
					Pattern pattern = null;
					pattern = java.util.regex.Pattern.compile(((ZipCode) annotation).pattern());
					Matcher matcher = pattern.matcher(obj == null ? "" : obj.toString());
					b = matcher.matches();
					log.debug("�����ʱ�:" + b);
				}
				if (annotation instanceof Date)
				{

					Object obj = field_new.get(t);
					java.util.Date date = ((java.util.Date) obj);
					String obj_str = DateFromat.formatDate(date, "yyyy-MM-dd");
					Pattern pattern = null;
					pattern = java.util.regex.Pattern.compile(((org.mec.validation.annotation.Date) annotation).pattern());
					Matcher matcher = pattern.matcher(obj_str == null ? "" : obj_str);
					b = matcher.matches();
					log.debug("��������:" + b);
				}
			}

		} catch (SecurityException e)
		{
			e.printStackTrace();
		} catch (IllegalArgumentException e)
		{
			e.printStackTrace();
		} catch (IllegalAccessException e)
		{
			e.printStackTrace();
		} catch (NoSuchFieldException e)
		{
			e.printStackTrace();
		}
		return b;
	}

	public static void main(String[] args) {
		PatternCompile pc=new PatternCompile();
		System.out.println("Can not set java.lang. field com.itsv.olderpeople.personinfo.dto.PersonInfoDto.age to java.lang.String".indexOf("Integer")!=-1);
		Student st=new Student();
		st.setPhone("13910204456");
		System.out.println(pc.compile(st,"phone"));
	}
}
