package org.epe.core.emnu;

/**
 * 文件护展名,支持xls,xlsx。即office2003与office2007
 * 
 * @author 李声波
 * 
 */
public enum FileExtendsName
{

	xls("xls"), xlsx("xlsx");

	private String fileExtendsName;

	private FileExtendsName(String fileExtendsName)
	{
		this.fileExtendsName = fileExtendsName;
	}

	public String getFileExtendsName()
	{
		return fileExtendsName;
	}
}
