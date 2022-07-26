import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;

import java.util.LinkedHashMap;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDriven_Java {

	public LinkedHashMap<String, String> readExcel(String directory,String Flag)
	{
		XSSFWorkbook workbook;
		XSSFSheet sheet;
		XSSFRow row;
		LinkedHashMap<String, String> data=new LinkedHashMap<String, String>();
		ArrayList firstrowvalues;
		
		File file=new File(directory);
		try {
			FileInputStream fis=new FileInputStream(file);
			workbook=new XSSFWorkbook(fis);
			sheet=workbook.getSheetAt(0);
			firstrowvalues=new ArrayList();
			row = sheet.getRow(0);
			int totalrows=row.getLastCellNum();
			//For Fetching the 1st row from the sheet
			for(int head=0;head<totalrows;head++)
			{
				String headvalue=row.getCell(head).getStringCellValue();
				firstrowvalues.add(headvalue);
			}
			
			//To get the data from each column
			for(int head=0;head<totalrows;head++)
			{
				//To check column by column
				//int icell=row.getCell(head).getColumnIndex();
				int rowscount=sheet.getLastRowNum();
				for(int row1=1;row1<rowscount;row1++)
				{
					//fetches the Testcase_ID column,iteration starts from row 2
					String cellname=sheet.getRow(row1).getCell(0).getStringCellValue();
					//Validates if any value in the column matches with the flag
					if(cellname.equalsIgnoreCase(Flag))
					{
						for(int head1=0;head1<row.getLastCellNum();head1++)
						{
							String headvalue1=row.getCell(head1).getStringCellValue();
							if(headvalue1.equals(firstrowvalues.get(head1)))
							{
								String head2=firstrowvalues.get(head1).toString();
								String Values=sheet.getRow(row1).getCell(head1).toString();
								data.put(head2, Values);
							}
						}
						return data;
					}
				}
			}
			System.out.println(data.toString());
		}
		catch (Exception e) {
			System.out.println(e.getMessage());
		}
		return data;
	}
	
	public String getData(LinkedHashMap<String, String>hs , String column) {
		if(hs.get(column).isEmpty())
		{
			return null;
		}
		
		return hs.get(column).toString();
		
	}
	
	public static void main(String[] args) {
		String excel_directory=System.getProperty("user.dir")+"\\src\\test\\java\\datatable\\DataDrivenDemo.xlsx";
        String testId_Flag="Scenario-2";
        DataDriven_Java runner=new DataDriven_Java();
        LinkedHashMap<String, String>hs =runner.readExcel(excel_directory, testId_Flag);
        System.out.println(hs);
        String ColumnValue="Zone 3";
        String Zone_3=runner.getData(hs,ColumnValue);
        System.out.println("TestData in "+ColumnValue+ " --> " +Zone_3);
	}

}
