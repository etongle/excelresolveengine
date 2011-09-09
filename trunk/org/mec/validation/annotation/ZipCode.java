package org.mec.validation.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * �ʱ���֤
 * 
 * @author ������
 * 
 */
@Documented
@Inherited
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ZipCode {
	/**
	 * ��֤�ʱ�
	 * @return
	 */
	public String pattern() default "[1-9]{1}(\\d+){5}";
}
