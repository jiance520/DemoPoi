package com.test;

import java.util.*;
import java.io.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;

// 读 xlsx
public class Test2 {

	public static void main(String[] args) {
		try {
			FileInputStream in = new FileInputStream("C:\\Ddisk\\1.xlsx");
			
			Workbook wk = WorkbookFactory.create(in); // 读文件  
			
			Sheet s1 = wk.getSheetAt(0); // 得到 sheet
			
			int lastnum = s1.getLastRowNum(); // 得到最后1行的下标
			System.out.println(" lastnum :"+lastnum);
			// 取行
			for(int i=0;i<=lastnum;i++){
				Row row = s1.getRow(i); 
				//取列
				int firstcell = row.getFirstCellNum();
				//System.out.println(" firstcell = "+firstcell);
				int lastcell = row.getLastCellNum();
				//System.out.println(" lastcell = "+lastcell);
				for(int j=firstcell;j<lastcell;j++){
					Cell cell = row.getCell(j);
					String vs = getCellValue(cell);// 拿值
					
					System.out.println(" value = "+vs);
				}
			}
			
			
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// 类型转化
	private static String getCellValue(Cell cell){
		String vs = null;
		
		if(cell.getCellType() == Cell.CELL_TYPE_BLANK){
			vs = "" ;
		}
		if(cell.getCellType() == Cell.CELL_TYPE_BOOLEAN){
			vs = String.valueOf(cell.getBooleanCellValue());
		}
		if(cell.getCellType() == Cell.CELL_TYPE_ERROR){
			vs = String.valueOf(cell.getErrorCellValue());
		}
		if(cell.getCellType() == Cell.CELL_TYPE_FORMULA){
			vs = String.valueOf(cell.getCellFormula());
		}
		if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC){
			vs = String.valueOf(cell.getNumericCellValue());
		}
		if(cell.getCellType() == Cell.CELL_TYPE_STRING){
			vs = cell.getStringCellValue();
		}
		
		return vs;
	}
}
