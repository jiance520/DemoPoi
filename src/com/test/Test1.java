package com.test;

import java.io.*;
import java.util.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;

//д xlsx
public class Test1 {

	public static void main(String[] args) {
		
		Workbook wk = new HSSFWorkbook();
		
		try {
			FileOutputStream  out = new FileOutputStream("C:\\Ddisk\\1.xlsx");
			//����sheet
			Sheet s1 = wk.createSheet("poi����1");
			
			//��ʽ����
			CellStyle  cs = wk.createCellStyle();
			
			//���嶨��
			Font ft = wk.createFont();
			     // ������ɫ
			     ft.setColor(HSSFColor.BLUE.index);
			     // �����С
			     ft.setFontHeightInPoints(Short.valueOf("24"));
			     // ����
			     ft.setBoldweight(Font.BOLDWEIGHT_BOLD);
			     // д��
			     ft.setFontName("����");
			
			// ������� ��ʽ
			cs.setFont(ft);
			
			// ������ɫ
			cs.setFillForegroundColor(IndexedColors.YELLOW.index);
			cs.setFillPattern(CellStyle.SOLID_FOREGROUND);
			
			//����������
			cs.setBorderBottom(CellStyle.BORDER_DASHED);
			
			//������ row;�±�� 0 ��ʼ
			Row r1 = s1.createRow(0); // ��1��
			//����cell;�±�� 0 ��ʼ
			Cell c1 = r1.createCell(0); // 1 �� 1��
			     c1.setCellStyle(cs); // ����ʽ
			//cell ��ֵ
			     c1.setCellValue("����1");
			Cell c2 = r1.createCell(1); // 1 �� 2��
			     c2.setCellValue("����2"); 
			     
			Row r2 = s1.createRow(1); // ��2��   
			Cell c3 = r2.createCell(0); // 2 �� 1��
		     c3.setCellValue("����3"); 
		    Cell c4 = r2.createCell(1); // 2 �� 2��
		     c4.setCellValue("����4"); 
			
		    // ���� �� 1 �� �Զ���չ  ; �±� �� �� ���±�
		    s1.autoSizeColumn(0);
		     
			wk.write(out);
		} catch (Exception e) {
			e.printStackTrace();
		}

		System.out.println("--------------");
	}

}
