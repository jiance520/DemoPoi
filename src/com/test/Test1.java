package com.test;

import java.io.*;
import java.util.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;

//写 xlsx
public class Test1 {

	public static void main(String[] args) {
		
		Workbook wk = new HSSFWorkbook();
		
		try {
			FileOutputStream  out = new FileOutputStream("C:\\Ddisk\\1.xlsx");
			//创建sheet
			Sheet s1 = wk.createSheet("poi测试1");
			
			//样式定义
			CellStyle  cs = wk.createCellStyle();
			
			//字体定义
			Font ft = wk.createFont();
			     // 字体颜色
			     ft.setColor(HSSFColor.BLUE.index);
			     // 字体大小
			     ft.setFontHeightInPoints(Short.valueOf("24"));
			     // 粗体
			     ft.setBoldweight(Font.BOLDWEIGHT_BOLD);
			     // 写法
			     ft.setFontName("楷体");
			
			// 字体加入 样式
			cs.setFont(ft);
			
			// 背景颜色
			cs.setFillForegroundColor(IndexedColors.YELLOW.index);
			cs.setFillPattern(CellStyle.SOLID_FOREGROUND);
			
			//设置下虚线
			cs.setBorderBottom(CellStyle.BORDER_DASHED);
			
			//创建行 row;下标从 0 开始
			Row r1 = s1.createRow(0); // 第1行
			//创建cell;下标从 0 开始
			Cell c1 = r1.createCell(0); // 1 行 1列
			     c1.setCellStyle(cs); // 加样式
			//cell 放值
			     c1.setCellValue("测试1");
			Cell c2 = r1.createCell(1); // 1 行 2列
			     c2.setCellValue("测试2"); 
			     
			Row r2 = s1.createRow(1); // 第2行   
			Cell c3 = r2.createCell(0); // 2 行 1列
		     c3.setCellValue("测试3"); 
		    Cell c4 = r2.createCell(1); // 2 行 2列
		     c4.setCellValue("测试4"); 
			
		    // 设置 第 1 行 自动扩展  ; 下标 是 行 的下标
		    s1.autoSizeColumn(0);
		     
			wk.write(out);
		} catch (Exception e) {
			e.printStackTrace();
		}

		System.out.println("--------------");
	}

}
