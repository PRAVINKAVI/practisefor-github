package sql_sever_connection;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.*;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Sql_sever_connection_first_attempt {

	public static String Sqlquary() {
		Scanner sc = new Scanner(System.in);
		String quarry = sc.nextLine();

		return quarry;
	}

	public static void main(String[] args) throws ClassNotFoundException, SQLException, IOException {

		Class.forName("oracle.jdbc.OracleDriver");

		Connection con = DriverManager.getConnection("jdbc:oracle:thin:@localhost:1521:xe", "hr", "admin");

		Statement steat = con.createStatement();

		ResultSet st = steat.executeQuery(Sqlquary());

		while (st.next()) {

			String str = (st.getString(1) + " " + st.getString(2) + " " + st.getString(3));
			
			File f=new File("C:\\Users\\Kaliraj\\OneDrive\\Desktop\\pravinraj\\practising file\\practise 6.xlsx");
			if(!f.exists())
			{
			f.createNewFile();
			}
			FileOutputStream fos = new FileOutputStream(f);
			XSSFWorkbook workbook = new XSSFWorkbook();
			XSSFSheet sheet = workbook.createSheet("pravinraj");
			XSSFRow row = sheet.createRow(5);
			XSSFCell cell = row.createCell(5);

			int rownum = sheet.getLastRowNum();
			int cellnum = sheet.getRow(rownum).getLastCellNum();
			System.out.println(rownum);
			System.out.println(cellnum);

			for (int i = 0; i <=2 ; i++) {
				for (int j = 0; j <=1 ; j++) {
					cell.setCellValue(st.getString(3));
					
					workbook.write(fos);
					
				}
			}

			fos.close();
		}

		con.close();

	}

}
