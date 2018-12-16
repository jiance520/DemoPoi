package com.action;

import java.io.*;

import com.entity.*;

import java.util.*;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;

public class ExcelServlet extends HttpServlet {

	/**
	 * Constructor of the object.
	 */
	public ExcelServlet() {
		super();
	}

	/**
	 * Destruction of the servlet. <br>
	 */
	public void destroy() {
		super.destroy(); // Just puts "destroy" string in log
		// Put your code here
	}

	/**
	 * The doGet method of the servlet. <br>
	 *
	 * This method is called when a form has its tag value method equals to get.
	 * 
	 * @param request the request send by the client to the server
	 * @param response the response send by the server to the client
	 * @throws ServletException if an error occurred
	 * @throws IOException if an error occurred
	 */
	public void doGet(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {

		doPost(request, response);
	}

	/**
	 * The doPost method of the servlet. <br>
	 *
	 * This method is called when a form has its tag value method equals to post.
	 * 
	 * @param request the request send by the client to the server
	 * @param response the response send by the server to the client
	 * @throws ServletException if an error occurred
	 * @throws IOException if an error occurred
	 */
	public void doPost(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {

		        String filename = "student.xlsx";
		        
		        List<Student> list = new ArrayList<Student>();
		                      list.add(new Student(1, "tom1", "man", 18));
		                      list.add(new Student(2, "tom2", "man", 18));
		                      list.add(new Student(3, "tom3", "man", 18));
		                      list.add(new Student(4, "tom4", "man", 18));
		        
		        //将集合生成xlsx让客户端下载
				Workbook wb = getExcel(list); 
				
				response.setContentType("application/vnd.ms-excel");    
				response.setHeader("Content-disposition", "attachment;filename="+filename);    
				
				OutputStream ouputStream = response.getOutputStream();    
				wb.write(ouputStream);    
				ouputStream.flush();    
				ouputStream.close(); 
	}

	private Workbook getExcel(List<Student> list){
		Workbook wb = new HSSFWorkbook(); 
		try {
			//创建表格
			Sheet sheet = wb.createSheet("学生01");
			
			//创建表头-------------------------------------
			Row row = sheet.createRow(0);
			//设置行高
			row.setHeightInPoints(30);
			
			//创建样式
			CellStyle cs = wb.createCellStyle();
			cs.setAlignment(CellStyle.ALIGN_CENTER);// 左右居中
			cs.setVerticalAlignment(CellStyle.VERTICAL_CENTER);//上下居中
			
			//创建字体
			Font ft = wb.createFont();
			     ft.setBoldweight(Font.BOLDWEIGHT_BOLD);// 粗体
			     ft.setFontHeightInPoints((short)18); // 字体大小
			     ft.setFontName("楷体"); // 写法
			     
			cs.setFont(ft);
			  
			//设置背景
			cs.setFillForegroundColor(IndexedColors.YELLOW.index);
	        cs.setFillPattern(CellStyle.SOLID_FOREGROUND);
			
			//创建单元格
			Cell ch = row.createCell(0);
			//设置单元格样式
			ch.setCellStyle(cs);
			//创建单元格
			ch.setCellValue("学号");
			
			ch = row.createCell(1);
			ch.setCellStyle(cs);
			ch.setCellValue("姓名");
			
			ch = row.createCell(2);
			ch.setCellStyle(cs);
			ch.setCellValue("性别");
			
			ch = row.createCell(3);
			ch.setCellStyle(cs);
			ch.setCellValue("年龄");
			//----------------------------------------------
			
			//创建内容
			for(int i=0;i<list.size();i++){
				Student temp = list.get(i);
				Row row2 = sheet.createRow(i+1);// 要加 1，第1行是标题
				
				Cell cc = row2.createCell(0);
				cc.setCellValue(temp.getSid());
				
				cc = row2.createCell(1);
				cc.setCellValue(temp.getSname());
				
				cc = row2.createCell(2);
				cc.setCellValue(temp.getSex());
				
				cc = row2.createCell(3);
				cc.setCellValue(temp.getAge());
			}
			
			//自动调整第1行标题
			sheet.autoSizeColumn(0);	
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		return wb;
	}
	/**
	 * Initialization of the servlet. <br>
	 *
	 * @throws ServletException if an error occurs
	 */
	public void init() throws ServletException {
		// Put your code here
	}

}
