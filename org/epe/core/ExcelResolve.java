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
			if (headCell.getFieldType() == String.class && field.getType()==String.class) {
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
			e.printStackTrace();
			if (e.getMessage().indexOf("Integer") != -1) {
				throw new ResolveException(headCell.getName() + "�ֶβ�Ϊ�Ϸ�������");
			} else if (e.getMessage().indexOf("Date") != -1) {
				throw new ResolveException(headCell.getName() + "�ֶβ��ǺϷ������ڸ�ʽ");
			} else if (e.getMessage().indexOf("Float") != -1) {
				throw new ResolveException(headCell.getName() + "�ֶβ��ǺϷ���С������");
			}
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
					if (ecell != null  && headCell.getName().equals(ecell.name())) {
						
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
			//if (getSheetItems() != null) {
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
			//} else {
			//	genericExcel(workbook, getHeadCells(), getSheetName());
			//}

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
	//		List<Sheet> sheets = aer.getSheet("������ҵ��չ���");
	//		for (Sheet sheet : sheets) {
	//			Row row = null;
	//			Cell cell = null;
	//			for (int j = 0; j < 10; j++) {
	//				row = sheet.createRow(j);
	//				for (int j2 = 0; j2 < 32; j2++) {
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
	//			for (int i = 1; i < 6; i++) {
	//				row = sheet.createRow(i);
	//				for (int j = 0; j < 32; j++) {
	//					row.createCell(j).setCellStyle(cellStyle);
	//				}
	//			}
	//			/******��һ��******/
	//			CellRangeAddress cra00 = new CellRangeAddress(0, 0, 0, 1);
	//			sheet.addMergedRegion(cra00);
	//			CellRangeAddress cra02 = new CellRangeAddress(0, 0, 2, 30);
	//			sheet.addMergedRegion(cra02);
	//			sheet.getRow(0).getCell(0).setCellValue("��(��):ʯ��ɽ");
	//			sheet.getRow(0).getCell(31).setCellValue("��λ:��");
	//			/******�ڶ���******/
	//			CellRangeAddress cra10 = new CellRangeAddress(1, 1, 0, 31);
	//			sheet.addMergedRegion(cra10);
	//			sheet.getRow(1).getCell(0).setCellValue("2010��ȱ������������乤��������������(��)");
	//			/******������******/
	//			CellRangeAddress cra20 = new CellRangeAddress(2, 4, 0, 0);
	//			sheet.addMergedRegion(cra20);
	//			sheet.getRow(2).getCell(0).setCellValue("ָ������ ");
	//			CellRangeAddress cra21 = new CellRangeAddress(2, 3, 1, 5);
	//			sheet.addMergedRegion(cra21);
	//			sheet.getRow(2).getCell(1).setCellValue("�����˿ڷ�չ��� ");
	//			CellRangeAddress cra26 = new CellRangeAddress(2, 2, 6, 31);
	//			sheet.addMergedRegion(cra26);
	//			sheet.getRow(2).getCell(6).setCellValue("������ҵ��չ���");
	//			/******������******/
	//			CellRangeAddress cra36 = new CellRangeAddress(3, 3, 6, 9);
	//			sheet.addMergedRegion(cra36);
	//			sheet.getRow(3).getCell(6).setCellValue("����άȨ");
	//			CellRangeAddress cra310 = new CellRangeAddress(3, 3, 10, 11);
	//			sheet.addMergedRegion(cra310);
	//			sheet.getRow(3).getCell(10).setCellValue("���������ʩ");
	//			sheet.getRow(3).getCell(12).setCellValue("���긣��");
	//			CellRangeAddress cra313 = new CellRangeAddress(3, 3, 13, 17);
	//			sheet.addMergedRegion(cra313);
	//			sheet.getRow(3).getCell(13).setCellValue("����ҽ�ƻ������");
	//			CellRangeAddress cra318 = new CellRangeAddress(3, 3, 18, 29);
	//			sheet.addMergedRegion(cra318);
	//			sheet.getRow(3).getCell(18).setCellValue("����Ⱥ����֯");
	//			CellRangeAddress cra330 = new CellRangeAddress(3, 3, 30, 31);
	//			sheet.addMergedRegion(cra330);
	//			sheet.getRow(3).getCell(30).setCellValue("����Ⱥ����֯");
	//			/******������******/
	//			sheet.getRow(4).getCell(1).setCellValue("60�����������˿���");
	//			sheet.getRow(4).getCell(2).setCellValue("70�����������˿���");
	//			sheet.getRow(4).getCell(3).setCellValue("80�����������˿���");
	//			sheet.getRow(4).getCell(4).setCellValue("100�����������˿���");
	//			sheet.getRow(4).getCell(5).setCellValue("�������ͥ�˿���");
	//			sheet.getRow(4).getCell(6).setCellValue("����ϵͳ�Ӵ����ŷ��ʴ���");
	//			sheet.getRow(4).getCell(7).setCellValue("���귨��Ԯ������");
	//			sheet.getRow(4).getCell(8).setCellValue("���ϰ�����");
	//			sheet.getRow(4).getCell(9).setCellValue("άȨЭ����֯��");
	//			sheet.getRow(4).getCell(10).setCellValue("����վ/����/����");
	//			sheet.getRow(4).getCell(11).setCellValue("�����˲μ�����");
	//			sheet.getRow(4).getCell(12).setCellValue("���ܸ��䲹������������");
	//			sheet.getRow(4).getCell(13).setCellValue("����ҽԺ��");
	//			sheet.getRow(4).getCell(14).setCellValue("���д�λ�� ");
	//			sheet.getRow(4).getCell(15).setCellValue("���չػ�ҽԺ��");
	//			sheet.getRow(4).getCell(16).setCellValue("��λ��");
	//			sheet.getRow(4).getCell(17).setCellValue("�����Ժ����");
	//			sheet.getRow(4).getCell(18).setCellValue("����Э��");
	//			sheet.getRow(4).getCell(19).setCellValue("���вμ�����");
	//			sheet.getRow(4).getCell(20).setCellValue("�ӡ�������Э��");
	//			sheet.getRow(4).getCell(21).setCellValue("���вμ�����");
	//			sheet.getRow(4).getCell(22).setCellValue("�֡�������Э��");
	//			sheet.getRow(4).getCell(23).setCellValue("���вμ�����");
	//			sheet.getRow(4).getCell(24).setCellValue("�С�������Э��");
	//			sheet.getRow(4).getCell(25).setCellValue("���вμ�����");
	//			sheet.getRow(4).getCell(26).setCellValue("��������");
	//			sheet.getRow(4).getCell(27).setCellValue("��ҵͶ�뾭��");
	//			sheet.getRow(4).getCell(28).setCellValue("��������������֯");
	//			sheet.getRow(4).getCell(29).setCellValue("�μ�����");
	//			sheet.getRow(4).getCell(30).setCellValue("����ѧУ");
	//			sheet.getRow(4).getCell(31).setCellValue("��У����");
	//			/******������******/
	//			sheet.getRow(5).getCell(0).setCellValue("��λ");
	//			sheet.getRow(5).getCell(1).setCellValue("��");
	//			sheet.getRow(5).getCell(2).setCellValue("��");
	//			sheet.getRow(5).getCell(3).setCellValue("��");
	//			sheet.getRow(5).getCell(4).setCellValue("��");
	//			sheet.getRow(5).getCell(5).setCellValue("��");
	//			sheet.getRow(5).getCell(6).setCellValue("��");
	//			sheet.getRow(5).getCell(7).setCellValue("��");
	//			sheet.getRow(5).getCell(8).setCellValue("��");
	//			sheet.getRow(5).getCell(9).setCellValue("��");
	//			sheet.getRow(5).getCell(10).setCellValue("��");
	//			sheet.getRow(5).getCell(11).setCellValue("��");
	//			sheet.getRow(5).getCell(12).setCellValue("��");
	//			sheet.getRow(5).getCell(13).setCellValue("��");
	//			sheet.getRow(5).getCell(14).setCellValue("�� ");
	//			sheet.getRow(5).getCell(15).setCellValue("��");
	//			sheet.getRow(5).getCell(16).setCellValue("��");
	//			sheet.getRow(5).getCell(17).setCellValue("��");
	//			sheet.getRow(5).getCell(18).setCellValue("��");
	//			sheet.getRow(5).getCell(19).setCellValue("��");
	//			sheet.getRow(5).getCell(20).setCellValue("��");
	//			sheet.getRow(5).getCell(21).setCellValue("��");
	//			sheet.getRow(5).getCell(22).setCellValue("��");
	//			sheet.getRow(5).getCell(23).setCellValue("��");
	//			sheet.getRow(5).getCell(24).setCellValue("��");
	//			sheet.getRow(5).getCell(25).setCellValue("��");
	//			sheet.getRow(5).getCell(26).setCellValue("��");
	//			sheet.getRow(5).getCell(27).setCellValue("��Ԫ");
	//			sheet.getRow(5).getCell(28).setCellValue("��");
	//			sheet.getRow(5).getCell(29).setCellValue("��");
	//			sheet.getRow(5).getCell(30).setCellValue("��");
	//			sheet.getRow(5).getCell(31).setCellValue("��");
	//			
	//
	//			aer.genericDataSource(serverPath);
	//		}
	//		System.out.println("ExcelResolve.main(�������)");
	//	}
}
