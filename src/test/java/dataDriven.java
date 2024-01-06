import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dataDriven {
	
	public ArrayList<String> getData(String testcaseName) throws IOException
	{
		FileInputStream fs = new FileInputStream("C:\\Users\\INDIA\\Downloads\\DemoData.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(fs);
		
		ArrayList<String> a = new ArrayList<String>();
		
		int sheets = wb.getNumberOfSheets();
		for(int i=0;i<sheets;i++)
		{
			if(wb.getSheetName(i).equalsIgnoreCase("testdata"))
			{
				XSSFSheet sheet =  wb.getSheetAt(i);
				
				//Identify Testcases column by scanning the entire 1st row
				
				Iterator<Row> rows = sheet.iterator(); //sheet is collection of rows
				Row firstrow = rows.next();
				Iterator<Cell> ce = firstrow.cellIterator(); //row is collection of cells
				
				int k=0,column=0;
				
				while(ce.hasNext())
				{
					Cell value = ce.next();
					if(value.getStringCellValue().equalsIgnoreCase("TestCases"))
					{
						column =  k;
					}
					k++;
				}
				System.out.println(column);
				
				//once column is identified then scan entire testcase column to identify purchase testcases
				
				while(rows.hasNext())
				{					
					Row r = rows.next();
					if(r.getCell(column).getStringCellValue().equalsIgnoreCase(testcaseName))
					{					
						//After you grab purchase testcase row pull all the data of that row and feed into text
						Iterator<Cell> cv = r.cellIterator();
						while(cv.hasNext())
						{
							Cell c = cv.next();
							if(c.getCellType() == CellType.STRING)
							{
								a.add(c.getStringCellValue());
							}
							else
							{
								a.add(NumberToTextConverter.toText(c.getNumericCellValue()));
								
							}
							
						}
					}
				}
				
			}
		}
		return a;
	}


	public static void main(String args[]) throws IOException
	{
		
	}
}
