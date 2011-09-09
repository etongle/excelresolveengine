package org.epe.core;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.epe.core.emnu.FileExtendsName;

/**
 * ��Ȩ˵����
 * 
	  Copyright(c) 2011 ERE1.2
	  
	 EREΪһ����Դ��Ŀ��Ŀǰ���ع���ERE�������ݣ����û�ģ�����ڽ����У�ERE��ѭGPLЭ�飬
	  �û����������޸ı����,���ڻ������ϴ����߷���ERE����ע��������
	  
	 ERE1.2֧��office2003��office2007,����POI3.8,ΪĿǰ���°汾��ERE1.2ֻ�ع������ж�ȡ��
	  ���ж�ȡ�븴�ӱ�ͷ�㷨��û������ع������ں�����ɡ�ERE1.2����������Ŀ�м򵥵ı��
	  ���뵼�����Լ���Sheet�ĵ��뵼������֧��VBA���������Excel����VBA��ERE�����޷���ȡ��
	  ��������������뽫VBAɾ���� ERE�ɹ�Ӧ����ɽ���ƶ��������ϵͳ���㽭�ƶ��������ƽ̨��
	  
	   �ı���־
	  1���ع����ж�ȡ 
	  2������ע�����ã�����Ҫ����EOM�����б�עע�����ӳ��ָ���ֶ� 
	  3��֧�ֶ�sheet
	  4������EOM��ERE1.2��ǰ�İ汾��֧��EOM����Excel�����ϵӳ�䣬���Է����ӳ�����
	  5�����ж�ȡ�븴�ӱ�ͷ��δ�ع������ں����汾�Ƴ���ERE1.2��ǰ�İ汾֧�֡�
	  6���ݲ�֧��Excel��ʽ
	  7��EDM��Excel���ݿ�ӳ�䣩��δ�ع���ERE1.2��ǰ�İ汾֧�֡�
	  
	  ERE���Ľ����࣬ʵ����EOM��Excel object Mapping��,OEM(Object Excel mapping)
	  �N����Ե���List<T>�е����ݣ�Ҳ�����н�Excel�е�������List<T>����ʽ���ء� list�еĶ�������û����ж��塣
	  ��Ŀ��ʵ���б�עECell���Խ�����ӳ�䵽Excel��
 * 
 * @author ������
 * 
 */
public abstract class AbstractExcelResolve
{

	protected Workbook[] workbook;
	protected Field[] fields;
	protected String serverPath;
	private InputStream is;

	private List<SheetItem> sheetItems;
	private List<HeadCell> headCells;
	private String sheetName;
	protected int fieldCount;
	
	protected List<ResolveException> exceptionlist; 

	/**
	 * ��������Դ
	 * 
	 * @param serverPath
	 *            ������·��
	 * @return
	 */
	public void genericDataSource(String serverPath)
	{
		for (int i = 0; i < this.workbook.length; i++)
		{
			Workbook wb = this.workbook[i];
			String filename = serverPath;
			if (wb instanceof XSSFWorkbook)
			{
				filename = filename + "x";
			}

			FileOutputStream out;
			try
			{
				out = new FileOutputStream(filename);
				wb.write(out);
				out.flush();
				out.close();
				// return true;
			} catch (FileNotFoundException e)
			{
				e.printStackTrace();
			} catch (IOException e)
			{
				e.printStackTrace();
			}
		}

	}

	/**
	 * ��ȡ����Դ
	 * 
	 * @param serverPath
	 */
	public void readerDataSource()
	{
		try
		{
			is = new FileInputStream(serverPath);
		} catch (FileNotFoundException e)
		{
			e.printStackTrace();
		}
	}

	/**
	 * ��ȡ����Դ
	 * 
	 * @param serverPath
	 */
	public void readerDataSource(InputStream inputStream)
	{
		try
		{
			is = inputStream;
		} catch (Exception e)
		{
			e.printStackTrace();
		}
	}

	/**
	 * ����������
	 */
	protected void createWorkbook()
	{
		try
		{
			if (is != null)
			{
				String extendsName = getServerPath().substring(getServerPath().lastIndexOf(".") + 1);
				if (extendsName.equals(FileExtendsName.xls.getFileExtendsName()))
					this.workbook = new Workbook[] { new HSSFWorkbook(is) };
				if (extendsName.equals(FileExtendsName.xlsx.getFileExtendsName()))
					this.workbook = new Workbook[] { new XSSFWorkbook(is) };
			} else
			{
				this.workbook = new Workbook[] { new HSSFWorkbook(), new XSSFWorkbook() };
			}
		} catch (FileNotFoundException e)
		{
			e.printStackTrace();
		} catch (IOException e)
		{
			e.printStackTrace();
		}
	}

	/**
	 * �������ݣ��÷������������List<T>�еĶ������ݣ�EREĿǰֻ֧��list
	 * 
	 * @param obj
	 *            List<T>
	 * @return List<HeadCell>
	 *         HeadCell��EREר�õ�Excel��Object�����ݴ洢�������е�EOM����ͨ���˶���洢����
	 */
	public abstract List<HeadCell> dataProcessFactory(Object obj);

	/**
	 * ����Excel���������List<SheetItem>���أ�SheetItem�洢�˶��sheet�е�����
	 * 
	 * @return List<SheetItem>
	 */
	public abstract <T extends Object> List<SheetItem> inputExcel();

	/**
	 * ����Excel
	 * 
	 * @param list
	 */
	public abstract void exportExcel(SheetItem... sheetItems);
	
	
	/**
	 * excel object ӳ�䷽��
	 * 
	 * @return
	 */
	public abstract <T extends Object> Map<String,List<T>> excelObjectMapping(List<HeadCell> list,
			Class<T> clss);
	
	/**
	 * ��������
	 * @param clss
	 * @return
	 */
	public abstract <T extends Object>  boolean exportReport(Sheet sheet);
	/**
	 * ��ȡExcel��Sheet	
	 * @return
	 */
	public abstract List<Sheet> getSheet(String sheetName);
	
	/**
	 * ��������
	 * 
	 * @return ������·��
	 */
	public String getServerPath()
	{
		return serverPath;
	}

	/**
	 * ������·��
	 * 
	 * @param serverPath
	 */
	public void setServerPath(String serverPath)
	{
		this.serverPath = serverPath;
	}
	
	/**
	 * �õ�sheet����
	 * 
	 * @return
	 */
	public List<SheetItem> getSheetItems()
	{
		return sheetItems;
	}

	/**
	 * ����sheet����
	 * 
	 * @param sheetItems
	 */
	public void setSheetItems(SheetItem... sheetItems)
	{
		this.sheetItems = new ArrayList<SheetItem>();
		for (SheetItem sheetItem : sheetItems)
		{
			this.sheetItems.add(sheetItem);
		}
	}

	/**
	 * �õ����ݼ�
	 * 
	 * @return
	 */
	public List<HeadCell> getHeadCells()
	{
		return headCells;
	}

	/**
	 * �������ݼ�
	 * 
	 * @param headCells
	 */
	public void setHeadCells(List<HeadCell> headCells)
	{
		this.headCells = headCells;
	}

	/**
	 * �õ�sheet���ƣ�ֻ�޵���sheetʹ��
	 * 
	 * @return
	 */
	public String getSheetName()
	{
		return sheetName;
	}

	/**
	 * ����sheet���ƣ�ֻ�޵���sheetʹ��
	 * 
	 * @param sheetName
	 */
	public void setSheetName(String sheetName)
	{
		this.sheetName = sheetName;
	}

	public Integer getFieldCount()
	{
		return fieldCount;
	}

	public void setFieldCount(Integer fieldCount)
	{
		this.fieldCount = fieldCount;
	}

	public List<ResolveException> getExceptionlist() {
		return exceptionlist;
	}

	public void setExceptionlist(List<ResolveException> exceptionlist) {
		this.exceptionlist = exceptionlist;
	}

}
