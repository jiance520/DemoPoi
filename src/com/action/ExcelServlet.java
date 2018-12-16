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
		        
		        //����������xlsx�ÿͻ�������
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
			//�������
			Sheet sheet = wb.createSheet("ѧ��01");
			
			//������ͷ-------------------------------------
			Row row = sheet.createRow(0);
			//�����и�
			row.setHeightInPoints(30);
			
			//������ʽ
			CellStyle cs = wb.createCellStyle();
			cs.setAlignment(CellStyle.ALIGN_CENTER);// ���Ҿ���
			cs.setVerticalAlignment(CellStyle.VERTICAL_CENTER);//���¾���
			
			//��������
			Font ft = wb.createFont();
			     ft.setBoldweight(Font.BOLDWEIGHT_BOLD);// ����
			     ft.setFontHeightInPoints((short)18); // �����С
			     ft.setFontName("����"); // д��
			     
			cs.setFont(ft);
			  
			//���ñ���
			cs.setFillForegroundColor(IndexedColors.YELLOW.index);
	        cs.setFillPattern(CellStyle.SOLID_FOREGROUND);
			
			//������Ԫ��
			Cell ch = row.createCell(0);
			//���õ�Ԫ����ʽ
			ch.setCellStyle(cs);
			//������Ԫ��
			ch.setCellValue("ѧ��");
			
			ch = row.createCell(1);
			ch.setCellStyle(cs);
			ch.setCellValue("����");
			
			ch = row.createCell(2);
			ch.setCellStyle(cs);
			ch.setCellValue("�Ա�");
			
			ch = row.createCell(3);
			ch.setCellStyle(cs);
			ch.setCellValue("����");
			//----------------------------------------------
			
			//��������
			for(int i=0;i<list.size();i++){
				Student temp = list.get(i);
				Row row2 = sheet.createRow(i+1);// Ҫ�� 1����1���Ǳ���
				
				Cell cc = row2.createCell(0);
				cc.setCellValue(temp.getSid());
				
				cc = row2.createCell(1);
				cc.setCellValue(temp.getSname());
				
				cc = row2.createCell(2);
				cc.setCellValue(temp.getSex());
				
				cc = row2.createCell(3);
				cc.setCellValue(temp.getAge());
			}
			
			//�Զ�������1�б���
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
