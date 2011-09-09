package org.mec.validation.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 身份证验证
 * @author 李声波
 *
 */
@Documented
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface IDCard {

	/**
	 * 验证类型，默认为18位验证
	 * @return
	 */
	public String new_pattern() default "^\\s*\\d{16}[\\dxX]{2}\\s*$";
	/**
	 * 验证类型，默认为15位验证
	 * @return
	 */
	public String pattern() default "^\\s*\\d{15}\\s*$";
}
