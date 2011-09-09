package org.epe.core;

import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.epe.annotations.ECell;

import com.itsv.gbp.core.util.DateFromat;

/**
 * ��Ȩ˵����
 * 
 * 
 * Copyright(c) 2011 ERE1.2 mec(������ũ΢���Ӽ����������)
 * 
 * EREΪһ����Դ��Ŀ��Ŀǰ���ع���ERE�������ݣ����û�ģ�����ڽ����У�ERE��ѭGPLЭ�飬
 * �û����������޸ı����,���ڻ������ϴ����߷���ERE����ע��������
 * 
 * ERE1.2֧��office2003��office2007,����POI3.8,ΪĿǰ���°汾��ERE1.2ֻ�ع������ж�ȡ��
 * ���ж�ȡ�븴�ӱ�ͷ�㷨��û������ع������ں�����ɡ�ERE1.2����������Ŀ�м򵥵ı��
 * ���뵼�����Լ���Sheet�ĵ��뵼������֧��VBA���������Excel����VBA��ERE�����޷���ȡ�� ��������������뽫VBAɾ����
 * ERE�ɹ�Ӧ����ɽ���ƶ��������ϵͳ���㽭�ƶ��������ƽ̨��
 * 
 * �ı���־
 * 
 * 1���ع����ж�ȡ 2������ע�����ã�����Ҫ����EOM�����б�עע�����ӳ��ָ���ֶ� 3��֧�ֶ�sheet
 * 4������EOM��ERE1.2��ǰ�İ汾��֧��EOM����Excel�����ϵӳ�䣬���Է����ӳ�����
 * 5�����ж�ȡ�븴�ӱ�ͷ��δ�ع������ں����汾�Ƴ���ERE1.2��ǰ�İ汾֧�֡� 6���ݲ�֧��Excel��ʽ
 * 7��EDM��Excel���ݿ�ӳ�䣩��δ�ع���ERE1.2��ǰ�İ汾֧�֡�
 * 
 * ��Ŀ����·�� http://code.google.com/p/excelresolveengine/downloads/list
 * 
 * ERE���Ľ����࣬ʵ����EOM��Excel object Mapping��,OEM(Object Excel mapping)
 * �N����Ե���List<T>�е����ݣ�Ҳ�����н�Excel�е�������List<T>����ʽ���ء� list�еĶ�������û����ж��塣
 * ��Ŀ��ʵ���б�עECell���Խ�����ӳ�䵽Excel��
 * 
 * @author ������
 * 
 */
public class ExcelResolve extends AbstractExcelResolve {

	public ExcelResolve() {
		super();
	}

	private short headBackgroundColor;

	/**
	 * ����ֵ
	 * 
	 * @param headCell
	 *            �洢���ݵĵ�Ԫ�����
	 * @param cell
	 *            ��Ԫ��
	 */
	private void setCellValues(HeadCell headCell, Cell cell) {
		if (headCell.getFieldType() == Integer.class) {
			cell.setCellValue(Integer.parseInt(headCell.getValue() == null ? "0" : headCell.getValue().toString()));
		}
		if (headCell.getFieldType() == String.class) {
			cell.setCellValue(headCell.getValue() == null ? "" : headCell.getValue().toString());
		}
		if (headCell.getFieldType() == Double.class) {
			cell.setCellValue(Double.parseDouble(headCell.getValue() == null ? "0" : headCell.getValue().toString()));
		}
		if (headCell.getFieldType() == Date.class) {
			if (headCell.getValue() == null) {
				cell.setCellValue((double) 0);
			} else {

				cell.setCellValue(DateFromat.formatDate((Date) headCell.getValue(), headCell.getPaten()));
			}

		}
		if (headCell.getFieldType() == Float.class) {
			cell.setCellValue(Float.parseFloat(headCell.getValue() == null ? "0" : headCell.getValue().toString()));
		}
		if (headCell.getFieldType() == int.class) {
			cell.setCellValue(Integer.parseInt(headCell.getValue() == null ? "0" : headCell.getValue().toString()));
		}
	}

	/**
	 * ִ��EOMʱ��ֵ�ķ���
	 * 
	 * @param field
	 * @param obj
	 * @param headCell
	 */
	private void setObjectValues(Field field, Object obj, HeadCell headCell) throws ResolveException {
		try {
			if (headCell.getFieldType() == Integer.class && field.getType() == Integer.class) {
				field.set(obj, Integer.parseInt(headCell.getValue() == null ? "" : headCell.getValue().toString()));
			}
			if (headCell.getFieldType() == String.class) {
				field.set(obj, headCell.getValue() == null ? "" : headCell.getValue().toString());
			}
			if (headCell.getFieldType() == Double.class && field.getType() == Double.class) {
				field.setDouble(obj, Double.parseDouble(headCell.getValue() == null ? "" : headCell.getValue().toString()));
			}
			if (headCell.getFieldType() == Date.class) {
				Date date = (Date) headCell.getValue();
				field.set(obj, date);
			}
			if (field.getType() == Float.class) {
				field.set(obj, new Float(headCell.getValue().toString()));
			}
			if (headCell.getFieldType() == int.class) {
				field.setInt(obj, Integer.parseInt(headCell.getValue() == null ? "" : headCell.getValue().toString()));
			}
		} catch (NumberFormatException e) {
			e.printStackTrace();
		} catch (IllegalArgumentException e) {
			if (e.getMessage().indexOf("Integer") != -1) {
				throw new ResolveException(headCell.getName() + "�ֶβ�Ϊ�Ϸ�������");
			} else if (e.getMessage().indexOf("Date") != -1) {
				throw new ResolveException(headCell.getName() + "�ֶβ��ǺϷ������ڸ�ʽ");
			} else if (e.getMessage().indexOf("Float") != -1) {
				throw new ResolveException(headCell.getName() + "�ֶβ��ǺϷ���С������");
			}
			e.printStackTrace();
		} catch (IllegalAccessException e) {
			e.printStackTrace();
		}
	}

	/**
	 * �õ�Ŀ�굥Ԫ�񣬸÷���������ECellע��������Ӧ��Excel�ֶ�
	 * 
	 * @param listObject
	 * @return
	 * @throws IllegalArgumentException
	 * @throws IllegalAccessException
	 */
	protected List<HeadCell> getTargetCells(List<?> listObject) throws IllegalArgumentException, IllegalAccessException {
		fieldCount = 0;
		List<HeadCell> list = new ArrayList<HeadCell>();
		if (listObject.size() < 0)
			return null;
		for (int i = 0; i < listObject.size(); i++) {
			Class<?> classObj = listObject.get(i).getClass();

			fields = classObj.getDeclaredFields();
			for (Field field : fields) {

				field.setAccessible(true);
				ECell ecell = field.getAnnotation(ECell.class);
				if (ecell == null) {
					continue;
				}
				HeadCell hc = new HeadCell();
				hc.setName(ecell.name());
				hc.setColumnName(field.getName());
				hc.setValue(field.get(listObject.get(i)));
				hc.setHidden(ecell.isHidden());
				hc.setFieldType(field.getType());
				list.add(hc);
				if (i == 0)
					fieldCount++;
			}
		}
		return list;
	}

	protected <T extends Object> T getTargetObject(Class<T> clss) throws IllegalArgumentException, IllegalAccessException {
		Object obj_instance = null;
		Constructor[] constructors = clss.getDeclaredConstructors();
		try {
			for (int i = 0; i < constructors.length; i++) {
				constructors[i].setAccessible(true);
				obj_instance = constructors[i].newInstance(new Object[0]);
			}
		} catch (IllegalArgumentException e1) {
			e1.printStackTrace();
		} catch (InstantiationException e1) {
			e1.printStackTrace();
		} catch (IllegalAccessException e1) {
			e1.printStackTrace();
		} catch (InvocationTargetException e1) {
			e1.printStackTrace();
		}
		return (T) obj_instance;
	}

	/**
	 * �õ���ͷ����ɫ
	 * 
	 * @return
	 */
	public short getHeadBackgroundColor() {
		return headBackgroundColor;
	}

	/**
	 * ���ñ�ͷ����ɫ
	 * 
	 * @param headBackgroundColor
	 */
	public void setHeadBackgroundColor(short headBackgroundColor) {
		this.headBackgroundColor = headBackgroundColor;
	}

	@Override
	public List<HeadCell> dataProcessFactory(Object obj) {

		// return listValue;
		return null;
	}

	@Override
	public <T extends Object> List<SheetItem> inputExcel() {
		List<SheetItem> listSheet = new ArrayList<SheetItem>();
		this.createWorkbook();
		List<String> columnNames = null;
		int cellCount = 0;
		for (Workbook workbook : this.workbook) {
			for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
				columnNames = new ArrayList<String>();
				Sheet sheet = workbook.getSheetAt(i);
				SheetItem si = new SheetItem();
				si.setName(sheet.getSheetName());
				List<HeadCell> headCells = new ArrayList<HeadCell>();
				for (int j = 0; j < sheet.getPhysicalNumberOfRows(); j++) {
					Row row = sheet.getRow(j);
					if (j == 0)
						cellCount = row.getPhysicalNumberOfCells();
					for (int h = 0; h < cellCount; h++) {
						String name = "";
						Cell cell = row.getCell(h);
						HeadCell headCell = new HeadCell();
						if (j == 0) {
							name = cell.getStringCellValue();
							columnNames.add(name);
							continue;
						}
						headCell.setName(columnNames.get(h));
						if (cell != null) {
							switch (cell.getCellType()) {
							case HSSFCell.CELL_TYPE_STRING:
								headCell.setFieldType(String.class);
								headCell.setValue(cell.getStringCellValue());
								break;
							case HSSFCell.CELL_TYPE_NUMERIC:

								if (HSSFDateUtil.isCellDateFormatted(cell)) {
									headCell.setFieldType(Date.class);
									headCell.setValue(cell.getDateCellValue());
								} else {
									double value = cell.getNumericCellValue();
									String str_value = new String(value + "");
									String float_value = str_value.substring(str_value.lastIndexOf(".") + 1);
									// �����ѧ������
									if (str_value.indexOf("E") != -1) {
										BigDecimal bd = new BigDecimal(str_value);
										headCell.setFieldType(String.class);
										headCell.setValue(bd.toPlainString());
									}
									// ������������
									else if ("0".equals(float_value)) {
										headCell.setFieldType(Integer.class);
										headCell.setValue(str_value.substring(0, str_value.lastIndexOf(".")));
									} else {
										headCell.setFieldType(Double.class);
										headCell.setValue(value);
									}

								}
								break;
							case HSSFCell.CELL_TYPE_BOOLEAN:
								headCell.setFieldType(Boolean.class);
								headCell.setValue(cell.getBooleanCellValue());

								break;

							}

						} else {
							headCell.setFieldType(String.class);
							headCell.setValue("");
						}
						headCells.add(headCell);
					}
				}
				si.setHeadCells(headCells);
				listSheet.add(si);
			}
		}
		return listSheet;
	}

	/**
	 * excel object ӳ�䷽��
	 * 
	 * @return
	 */
	public <T extends Object> Map<String, List<T>> excelObjectMapping(List<HeadCell> list, Class<T> clss) {
		Map<String, List<T>> map = new HashMap<String, List<T>>();
		List<T> listMessage = new ArrayList<T>();
		List<T> list_new = new ArrayList<T>();
		StringBuffer message = null;
		// ����ʼ�ֶ�����
		fieldCount = 0;
		Field[] fields_new = clss.getDeclaredFields();
		fields = new Field[fields_new.length];
		for (Field field : fields_new) {
			field.setAccessible(true);
			ECell ecell = field.getAnnotation(ECell.class);
			if (ecell == null)
				continue;
			else {
				fields[fieldCount] = field;
				fieldCount++;

			}
		}
		int x = 0;
		T obj = null;

		for (int i = 0; i < list.size(); i++) {
			for (int j = 0; j < fieldCount; j++) {
				if (x == 0 || x % fieldCount == 0) {
					try {
						obj = getTargetObject(clss);
						message = new StringBuffer();
					} catch (IllegalArgumentException e1) {
						e1.printStackTrace();
					} catch (IllegalAccessException e1) {
						e1.printStackTrace();
					}

				}
				if (x >= list.size()) {
					break;

				}
				// ��List<HeadCell>�ó�������Excel����
				HeadCell headCell = list.get(x);
				// ӳ�����

				ECell ecell = fields[j].getAnnotation(ECell.class);
				try {
					if (ecell != null && headCell.getName().equals(ecell.name())) {
						setObjectValues(fields[j], obj, headCell);
					}
				} catch (ResolveException e) {
					e.printStackTrace();
					//if ("message".equals(fields[j].getName())) {
					message.append(e.getMessage() + ",");
					//try {
					//fields[j].set(obj, message.toString());
					//} catch (IllegalArgumentException e1) {
					//	e1.printStackTrace();
					//	} catch (IllegalAccessException e1) {
					//	e1.printStackTrace();
					//	}

				}
				//}

				x++;
				if (x % fieldCount == 0) {
					list_new.add(obj);
					listMessage.add((T) message);
				}
			}
		}
		map.put("list", list_new);
		map.put("message", listMessage);
		return map;
	}

	/**
	 * ����EXCEL
	 */
	@Override
	public void exportExcel(SheetItem... sheetItems) {
		this.createWorkbook();
		for (Workbook workbook : this.workbook) {
			// ���sheet����ʽ
			if (getSheetItems() != null) {
				for (SheetItem sheetItem : sheetItems) {
					// Sheet s = wb.createSheet();
					List<HeadCell> listValue = null;
					if (sheetItem.getList().getClass() == ArrayList.class) {
						ArrayList<?> list = (ArrayList<?>) sheetItem.getList();
						try {
							listValue = getTargetCells(list);
						} catch (IllegalArgumentException e) {
							e.printStackTrace();
						} catch (IllegalAccessException e) {
							e.printStackTrace();
						}
					}
					genericExcel(workbook, listValue, sheetItem.getName());
				}
			} else {
				genericExcel(workbook, getHeadCells(), getSheetName());
			}

		}
		genericDataSource(this.getServerPath());
	}

	private void genericExcel(Workbook workbook, List<HeadCell> list, String sheetName) {

		Sheet sheet = null;
		if (sheetName == null || "".equals(sheetName)) {
			sheet = workbook.createSheet();
		} else {
			sheet = workbook.createSheet(sheetName);
		}
		Row headRow = sheet.createRow(0);
		CellStyle cs = workbook.createCellStyle();
		// cs.setFillForegroundColor(headBackgroundColor);
		// cs.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cs.setBorderBottom((short) 1);
		cs.setBorderLeft((short) 1);
		cs.setBorderRight((short) 1);
		cs.setBorderTop((short) 1);
		cs.setAlignment(CellStyle.ALIGN_LEFT);
		int x = 0;
		for (int j = 1; j <= list.size(); j++) {
			Row row_value = sheet.createRow(j);
			for (int i = 0; i < fieldCount; i++) {
				// �ڻ�ȡ����֮ǰ���ɱ�ͷ
				if (j == 1) {
					Cell cell = headRow.createCell(i);
					HeadCell hc = list.get(i);
					if (hc == null || list.get(x) == null || list.get(x).getName() == null)
						continue;
					else {
						cell.setCellValue(hc.getName());
						cell.setCellStyle(cs);
						//�ϲ���Ԫ��
						//						CellRangeAddress cra=new CellRangeAddress(2, 2, 0, 15);
						//					    sheet.addMergedRegion(cra);
						sheet.setColumnWidth(i, list.get(x).getName() == null ? 1000 : (list.get(x).getName().length() > (list.get(x).getValue() == null ? 0 : list.get(x).getValue().toString()
								.length()) ? list.get(x).getName().length() : list.get(x).getValue().toString().length()) * 500);
					}
				}
				// ���ܴ���list���ϴ�С����������ѭ��
				if (x >= list.size()) {
					break;
				} else {
					// ��������
					Cell cell_value = row_value.createCell(i);
					cell_value.setCellStyle(cs);
					setCellValues(list.get(x), cell_value);
				}
				// ����ȡ�����ݵ�����
				++x;
			}
		}
	}

	@Override
	public <T> boolean exportReport(Sheet sheet) {
		return false;
	}

	@Override
	public List<Sheet> getSheet(String sheetName) {
		createWorkbook();
		List<Sheet> sheets = new ArrayList<Sheet>();
		for (Workbook workbk : this.workbook) {
			Sheet sheet = workbk.createSheet(sheetName);
			sheets.add(sheet);
		}
		return sheets;
	}
	/**
	 * ���Ե���EXCEl�ļ���ʽ
	 */
	//	public static void main(String[] args) {
	//		String serverPath = "c:\\test.xls";
	//
	//		AbstractExcelResolve aer = new ExcelResolve();
	//		aer.setServerPath(serverPath);
	//		List<Sheet> sheets = aer.getSheet("���乤�����籨��");
	//		for (Sheet sheet : sheets) {
	//			Row row = null;
	//			Cell cell = null;
	//			for (int j = 0; j < 8; j++) {
	//				row = sheet.createRow(j);
	//				for (int j2 = 0; j2 < 7; j2++) {
	//					cell = row.createCell(j2);
	//					cell.setCellValue(j2);
	//				}
	//			}
	//			/*********************���ø�ʽ**************************/
	//			CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
	//
	//			//���뷽ʽ
	//			cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
	//			//�Զ�����
	//			cellStyle.setWrapText(true);
	//			//�߿���ɫ
	//			cellStyle.setBorderBottom((short) 1);
	//			cellStyle.setBorderLeft((short) 1);
	//			cellStyle.setBorderRight((short) 1);
	//			cellStyle.setBorderTop((short) 1);
	//			//������ɫ
	//			cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
	//			cellStyle.setFillForegroundColor(HSSFColor.TURQUOISE.index);
	//			for (int i = 1; i < 7; i++) {
	//				row = sheet.createRow(i);
	//				for (int j = 0; j < 7; j++) {
	//					row.createCell(j).setCellStyle(cellStyle);
	//				}
	//			}
	//			/******��һ��******/
	//			CellRangeAddress cra00 = new CellRangeAddress(0, 0, 0, 1);
	//			sheet.addMergedRegion(cra00);
	//			CellRangeAddress cra02 = new CellRangeAddress(0, 0, 2, 4);
	//			sheet.addMergedRegion(cra02);
	//			sheet.getRow(0).getCell(0).setCellValue("��(��):ʯ��ɽ");
	//			sheet.getRow(0).getCell(6).setCellValue("��λ:��");
	//			/******�ڶ���******/
	//			CellRangeAddress cra10 = new CellRangeAddress(1, 1, 0, 6);
	//			sheet.addMergedRegion(cra10);
	//			sheet.getRow(1).getCell(0).setCellValue("2011��ȱ������������乤��������������(��)");
	//			/******������******/
	//			CellRangeAddress cra20 = new CellRangeAddress(2, 2, 0, 6);
	//			sheet.addMergedRegion(cra20);
	//			sheet.getRow(2).getCell(0).setCellValue("���g����������ʩ��� ");
	//			/******������******/
	//			CellRangeAddress cra30 = new CellRangeAddress(3, 3, 0, 6);
	//			sheet.addMergedRegion(cra30);
	//			sheet.getRow(3).getCell(0).setCellValue("ʵʩ\"�ǹ�ƻ�\"��������긣��������ʩ ");
	//			/******������******/
	//			CellRangeAddress cra40 = new CellRangeAddress(4, 4, 0, 3);
	//			sheet.addMergedRegion(cra40);
	//			sheet.getRow(4).getCell(0).setCellValue("����");
	//			CellRangeAddress cra44 = new CellRangeAddress(4, 4, 4, 6);
	//			sheet.addMergedRegion(cra44);
	//			sheet.getRow(4).getCell(4).setCellValue("ʹ����� ");
	//			/******������******/
	//			CellRangeAddress cra50 = new CellRangeAddress(5, 5, 0, 1);
	//			sheet.addMergedRegion(cra50);
	//			sheet.getRow(5).getCell(0).setCellValue("���������ǹ�����֮��");
	//			CellRangeAddress cra52 = new CellRangeAddress(5, 5, 2, 3);
	//			sheet.addMergedRegion(cra52);
	//			sheet.getRow(5).getCell(2).setCellValue("ũ�����긣��������ʩ");
	//			sheet.getRow(5).getCell(4).setCellValue("������ת");
	//			sheet.getRow(5).getCell(5).setCellValue("�ı���;");
	//			sheet.getRow(5).getCell(6).setCellValue("����");
	//			/******�ڰ���******/
	//			CellRangeAddress cra70 = new CellRangeAddress(7, 7, 0, 1);
	//			sheet.addMergedRegion(cra70);
	//			sheet.getRow(7).getCell(0).setCellValue("������:");
	//			CellRangeAddress cra72 = new CellRangeAddress(7, 7, 2, 3);
	//			sheet.addMergedRegion(cra72);
	//			sheet.getRow(7).getCell(2).setCellValue("�����:");
	//			CellRangeAddress cra75 = new CellRangeAddress(7, 7, 4, 7);
	//			sheet.addMergedRegion(cra75);
	//			sheet.getRow(7).getCell(5).setCellValue("�����:");
	//
	//			aer.genericDataSource(serverPath);
	//
	//		}
	//		System.out.println("ExcelResolve.main(�������)");
	//	}
}
