package main;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DbToExcel {
	public static void main(String[] args) throws ClassNotFoundException,
			IOException, SQLException {
		Class.forName("org.sqlite.JDBC");
		DbToExcel db = new DbToExcel();

		db.readFromDbWriteToExcel();

	}

	public void readFromDbWriteToExcel() throws IOException,
			ClassNotFoundException, SQLException {

		File[] folder = new File("D:\\Perindatabases").listFiles();
		for (File file : folder) {
			XSSFWorkbook workbook =  new XSSFWorkbook();

			System.out.println(file.getName());
			Connection connection = DriverManager
					.getConnection("jdbc:sqlite:D:\\Perindatabases\\"
							+ file.getName());

			Statement statement = connection.createStatement();
			String[] types = { "TABLE" };
			ArrayList<String> tableNames = new ArrayList<>();
			ResultSet resultSet = connection.getMetaData().getTables(file.getName(),null, "%", types);
			while (resultSet.next()) {
				tableNames.add(resultSet.getString(3).trim());
			}
		
			for (String table : tableNames) {
			System.out.println(table);
			
			XSSFSheet sheet = workbook.createSheet(table);
				
			 int count = sheet.getPhysicalNumberOfRows();
			 System.out.println("Fiziksel:"+sheet.getPhysicalNumberOfRows());
			   resultSet = statement.executeQuery("SELECT * FROM "+ table);
		      ResultSetMetaData rsmd = resultSet.getMetaData();
		      
			  int columnCount = rsmd.getColumnCount();
			  	ArrayList<Object> temp = new ArrayList<>();
			  	ArrayList<String> columnNames = new ArrayList<>();
				for (int i = 1; i < columnCount + 1; i++) {
				columnNames.add( rsmd.getColumnName(i));
				}
				
				while(resultSet.next()){
					for(int i=0;i<columnNames.size();i++){
						temp.add(resultSet.getString(columnNames.get(i)));
					}
					
				}
				
				int cellNumber = 0;
				XSSFRow row = sheet.createRow(count++);
				for (Object obj : temp) {		
					Cell cell = row.createCell(cellNumber++);

					if (obj instanceof Date)
						cell.setCellValue((Date) obj);
					else if (obj instanceof Boolean)
						cell.setCellValue((Boolean) obj);
					else if (obj instanceof String)
						cell.setCellValue((String) obj);
					else if (obj instanceof Double){
						cell.setCellValue((Double) obj);
					}
					
					
				}
				cellNumber=0;
			
			}
			String fileName = file.getName().substring(0,
					file.getName().indexOf('-'));
			File targetFolder = new File("D:\\excelDosyalari\\");
			if (!targetFolder.isDirectory())  new File("D:\\excelDosyalari\\").mkdir();
			
			String path = "D:\\excelDosyalari\\" + fileName + ".xls";
			File excelFile = new File(path);
		
			if (!file.exists())	file.createNewFile();
			FileOutputStream out = new FileOutputStream(excelFile);
			workbook.write(out);
			out.close();
		}
		
	}
}
