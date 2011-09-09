package org.mec.validation.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * �ֻ�������֤
 * 
 * @author ������
 * 
 */
@Documented
@Inherited
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface Phone{

	/**
	 * ��֤�ֻ�����
	 * @return
	 */
	public String pattern() default "^1[3|4|5|8][0-9]\\d{4,8}$";
}
