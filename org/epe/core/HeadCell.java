package org.epe.core;

import java.util.ArrayList;
import java.util.List;

/**
 * ERE专用Excel Object数据存储类
 * 
 * @author 李声波
 * 
 */
public class HeadCell
{
	public List<HeadCell> addHeadCell(HeadCell... headCell)
	{
		List<HeadCell> list = new ArrayList<HeadCell>();
		for (HeadCell headCell2 : headCell)
		{
			list.add(headCell2);
		}
		return list;
	}

	private String name;
	private String columnName;
	private Object value;
	private Class<?> fieldType;
	private boolean isHidden;
	private String paten = "yyyy-MM-dd";

	/**
	 * 得到列头中文名
	 * 
	 * @return
	 */
	public String getName()
	{
		return name;
	}

	/**
	 * 设置列头中文名
	 * 
	 * @param name
	 */
	public void setName(String name)
	{
		this.name = name;
	}

	/**
	 * 得到列头英文名
	 * 
	 * @return
	 */
	public String getColumnName()
	{
		return columnName;
	}

	/**
	 * 设置列头英文名
	 * 
	 * @param columnName
	 */
	public void setColumnName(String columnName)
	{
		this.columnName = columnName;
	}

	/**
	 * 得到值
	 * 
	 * @return
	 */
	public Object getValue()
	{
		return value;
	}

	/**
	 * 设置值
	 * 
	 * @param value
	 */
	public void setValue(Object value)
	{
		this.value = value;
	}

	/**
	 * 是否隐藏，暂未实现
	 * @return
	 */
	public boolean isHidden()
	{
		return isHidden;
	}

	/**
	 * 是否隐藏，暂未实现
	 * @return
	 */
	public void setHidden(boolean isHidden)
	{
		this.isHidden = isHidden;
	}

	/**
	 * 得到字段类型
	 * @return
	 */
	public Class<?> getFieldType()
	{
		return fieldType;
	}

	/**
	 * 设置字段类型
	 * @param fieldType
	 */
	public void setFieldType(Class<?> fieldType)
	{
		this.fieldType = fieldType;
	}

	public String getPaten()
	{
		return paten;
	}

	public void setPaten(String paten)
	{
		this.paten = paten;
	}

}
