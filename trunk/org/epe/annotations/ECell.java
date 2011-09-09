package org.epe.annotations;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel映射注解，可以将字段映射到Excel中
 * 
 * @author Administrator
 * 
 */
@Documented
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ECell
{

	/**
	 * 显示的名字
	 * 
	 * @return
	 */
	public String name();

	/**
	 * 是否隐藏
	 * 
	 * @return
	 */
	public boolean isHidden() default false;
}
