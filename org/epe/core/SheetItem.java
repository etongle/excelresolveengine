package org.epe.core;

import java.util.List;

/**
 * sheet集合对象,该类存储了所有的sheet中的数据
 * 
 * @author 李声波
 * 
 */
public class SheetItem {

	private String name;
	List<HeadCell> headCells;
	
	private List<?> list;

	/**
	 * 得到单个sheet中的数据
	 * @return
	 */
	public List<HeadCell> getHeadCells() {
		return headCells;
	}
	/**
	 * 设置单个sheet中的数据
	 */
	public void setHeadCells(List<HeadCell> headCells) {
		this.headCells = headCells;
	}

	/**
	 * sheet名称
	 * @return
	 */
	public String getName() {
		return name;
	}

	/**
	 * sheet名称
	 * @param name
	 */
	public void setName(String name) {
		this.name = name;
	}
	public List<?> getList()
	{
		return list;
	}
	public void setList(List<?> list)
	{
		this.list = list;
	}

}
