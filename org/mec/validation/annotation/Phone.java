package org.mec.validation.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 手机号码验证
 * 
 * @author 李声波
 * 
 */
@Documented
@Inherited
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface Phone{

	/**
	 * 验证手机号码
	 * @return
	 */
	public String pattern() default "^1[3|4|5|8][0-9]\\d{4,8}$";
}
