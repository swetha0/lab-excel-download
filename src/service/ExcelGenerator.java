package service;

import java.io.FileOutputStream;
import java.io.IOException;

import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import model.Prograd;

public class ExcelGenerator {
	String filename="C:\\Users\\SWETHA\\Desktop\\Book.xlsx";
	int i=1;
	FileOutputStream out;
	public Workbook excelGenerate(Prograd prograd, List<Prograd> list) throws IOException {
		try {
	
			Workbook hwb = new XSSFWorkbook();
			Sheet sheet=hwb.createSheet("ProGradDetails");
			Row row=sheet.createRow(0);
			
			row.createCell(0).setCellValue("ProGrad Name");
			row.createCell(1).setCellValue("ProGrad Id");
			row.createCell(2).setCellValue("ProGrad Rate");
			row.createCell(3).setCellValue("ProGrad Comment");
			row.createCell(4).setCellValue("ProGrad Recommend");
			
		 	
			for(Prograd fillSheet: list) {
	      	 
	      	  Row nextRows = sheet.createRow(i);
	      	nextRows.createCell(0).setCellValue(fillSheet.getName());
	      	nextRows.createCell(1).setCellValue(fillSheet.getId());
	      	nextRows.createCell(2).setCellValue(fillSheet.getRate());
	      	nextRows.createCell(3).setCellValue(fillSheet.getComment());
	      	nextRows.createCell(4).setCellValue(fillSheet.getRecommend());
	      	  
      	  }
			// Do not modify the lines given below
			out = new FileOutputStream(filename);
			hwb.write(out);
			return hwb;
			}
		catch (Exception e) {

				e.printStackTrace();
			}

		return null;
		
	}
}
