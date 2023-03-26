package sql_sever_connection;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excelpractise {

	public static void main(String[] args) throws IOException {
		
		
		FileInputStream fis=new FileInputStream("C:\\Users\\Kaliraj\\OneDrive\\Desktop\\pravinraj\\practising file\\practise 1.xlsx");
		FileOutputStream fos=new FileOutputStream("C:\\Users\\Kaliraj\\OneDrive\\Desktop\\pravinraj\\practising file\\practise 1.xlsx");
		XSSFWorkbook workbook=new XSSFWorkbook();
	XSSFSheet sheet=	workbook.createSheet("pravinraj");
	XSSFRow row=sheet.createRow(5);
	XSSFCell cell=sheet.createRow(5).createCell(5);
	
	int rownum=sheet.getLastRowNum();
	int cellnum=sheet.getRow(rownum).getLastCellNum();
	
	for(int i=0;i<=rownum;i++)
	{
		for(int j=0;j<=cellnum;j++)
		{
			cell.setCellValue(str);
		}
	}
	cell.setCellValue("kavibalan");
	}

}
