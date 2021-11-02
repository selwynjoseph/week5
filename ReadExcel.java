package week5.day2;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {

	public String[][] readData(String fileName) throws IOException {

		//setup workbook
		XSSFWorkbook wb=new XSSFWorkbook("./Data/"+fileName+".xlsx");
		
		//open sheet
		XSSFSheet ws=wb.getSheet("Sheet1");
		
		//get row and cell
		int rowcount=ws.getLastRowNum();
		int cellcount=ws.getRow(0).getLastCellNum();
		
		//declare a 2d array
		String[][] data=new String[rowcount][cellcount];
		
		for(int i=1; i<=rowcount; i++)
		{
			for(int j=0; j<cellcount; j++)
			{
				String text=ws.getRow(i).getCell(j).getStringCellValue();
				System.out.println(text);
				data[i-1][j]=text;
			}
		}
		
		//close workbook
		wb.close();
		return data;

	}

}
