package org.epe.core.test;

import java.util.List;

import org.epe.core.ExcelResolve;
import org.epe.core.ResolveException;
import org.epe.core.ResolveException;
import org.epe.core.SheetItem;

import com.itsv.olderpeople.personinfo.dto.PersonInfoDto;

public class Test {
	public static void main(String[] args) throws ResolveException {

//		List<Student> list = new ArrayList<Student>();
//		Student su1 = new Student();
//		 su1.setName("张三范德萨");
//		su1.setAge(18);
//		// su.setSex("男");
//		 su1.setSid(1);
//		list.add(su1);
//		for (int i = 0; i < 100; i++) {
//			Student su = new Student();
//			int id = i + 1;
//			su.setName("张三" + id);
//			su.setAge(1 + i);
//			su.setSex("男");
//			su.setSid(id);
//			list.add(su);
//		}

		ExcelResolve ope = new ExcelResolve();
		ope.setServerPath("c:\\老年人口信息Excel.xls");
//		List<HeadCell> headCells=ope.dataProcessFactory(list);
//		//单个sheet代码示例
//		ope.setHeadCells(headCells);
//		ope.exportExcel();
//		ope.sets
       // 多个sheet代码示例
//        SheetItem sheetItem=new SheetItem();
//        sheetItem.setName("学生1");
//        sheetItem.setHeadCells(headCells);
//        SheetItem sheetItem1=new SheetItem();
//        sheetItem1.setName("学生2");
//        sheetItem1.setHeadCells(headCells);
//        ope.setSheetItems(sheetItem,sheetItem1);
//        ope.exportExcel();
		  
        
//		//读取示例
//		ope.readerDataSource();
//		List<SheetItem> sheets=ope.inputExcel();
//		for (SheetItem sheetItem : sheets) {
//			System.out.println("sheet name:"+sheetItem.getName());
//			List<PersonInfoDto> list_student;
//			try {
//				list_student = ope.excelObjectMapping(sheetItem.getHeadCells(), PersonInfoDto.class);
//				for (PersonInfoDto personInfoDto : list_student) {
//					System.out.println("============="+personInfoDto.getName());
//				}
//			} catch (ResolveException e) {
//				// TODO Auto-generated catch block
//				e.printStackTrace();
//			}
//		}
	}
		
}
