package org.mec.validation.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 正整数验证
 * 
 * @author 李声波
 * 
 */
@Documented
@Inherited
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface PInteger {
	/**
	 * 验证正整数
	 * @return
	 */
	public String pattern() default "^[0-9]*[1-9][0-9]*$";
}
