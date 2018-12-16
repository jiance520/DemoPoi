package com.test;

import java.util.*;
import java.io.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;

// �� xlsx
public class Test2 {

	public static void main(String[] args) {
		try {
			FileInputStream in = new FileInputStream("C:\\Ddisk\\1.xlsx");
			
			Workbook wk = WorkbookFactory.create(in); // ���ļ�  
			
			Sheet s1 = wk.getSheetAt(0); // �õ� sheet
			
			int lastnum = s1.getLastRowNum(); // �õ����1�е��±�
			System.out.println(" lastnum :"+lastnum);
			// ȡ��
			for(int i=0;i<=lastnum;i++){
				Row row = s1.getRow(i); 
				//ȡ��
				int firstcell = row.getFirstCellNum();
				//System.out.println(" firstcell = "+firstcell);
				int lastcell = row.getLastCellNum();
				//System.out.println(" lastcell = "+lastcell);
				for(int j=firstcell;j<lastcell;j++){
					Cell cell = row.getCell(j);
					String vs = getCellValue(cell);// ��ֵ
					
					System.out.println(" value = "+vs);
				}
			}
			
			
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// ����ת��
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
