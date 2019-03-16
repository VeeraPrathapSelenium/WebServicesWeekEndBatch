package com.ReadExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;










public class ReadExcelData {
	
	
	public static XSSFWorkbook workbook;
	
	
	

	public static void main(String[] args) throws IOException {
		
		
		loadTestData();
		
		String data=getData("GenericData","Yes","Execute");
		
		System.out.println(data);

	}
	
	
	public static int searchColumn(String sheetname,String columnname)		
	{
		int colfound=0;
		
		
XSSFSheet sheet=workbook.getSheet(sheetname);
		
		XSSFRow row=sheet.getRow(0);
		
		XSSFCell cell;
		
	
		// total number of columns count	
			
			
		int totalcolcount=row.getPhysicalNumberOfCells()-1;
		
		for(int c=0;c<=totalcolcount;c++)
		{
			
			String actualvalue=sheet.getRow(0).getCell(c).getStringCellValue();
			
			if(columnname.equals(actualvalue))
			{
				
				colfound=c;
				break;
			}
			
			
		}
		
		
		return colfound;
		
		
	}
	
	
	
	public static int searchRowData(String sheetname,int colnumber,String rowdata)		
	{
		int rowfound=0;
		
		
XSSFSheet sheet=workbook.getSheet(sheetname);
		
			
		int totalrowcount=sheet.getPhysicalNumberOfRows()-1;
		rowcounter:
		for(int r=0;r<=totalrowcount;r++)
		{
			
			switch(sheet.getRow(r).getCell(colnumber).getCellTypeEnum())
			{
			
			
			case STRING:
				
				String actualvalue=sheet.getRow(r).getCell(colnumber).getStringCellValue();
				
				if(actualvalue.toLowerCase().equals(rowdata.toLowerCase()))
				{
					rowfound=r;
					break rowcounter;
				}
				
				
				
				break;
				
			case NUMERIC:
				int numvalue=(int) sheet.getRow(r).getCell(colnumber).getNumericCellValue();
				
				if(Integer.parseInt(rowdata)==numvalue)
				{
					rowfound=r;
					break rowcounter;
				}
				break;
			}
			
			
			
			
			
		}
		
		
		return rowfound;
		
		
	}
	
	
	
	
	
	public static String  getData(String sheetname,String rowdata,String columnname)		
	{
		String data="";
		
		int colnumber= searchColumn(sheetname,columnname);			
		System.out.println(colnumber);	
		
		
		
		int rownumber=searchRowData(sheetname,colnumber,rowdata);
		
		
		
XSSFSheet sheet=workbook.getSheet(sheetname);
		
			
		
			
			switch(sheet.getRow(rownumber).getCell(colnumber).getCellTypeEnum())
			{
			
			
			case STRING:
				
				data=sheet.getRow(rownumber).getCell(colnumber).getStringCellValue();
								
				break;
				
			case NUMERIC:
				int numvalue=(int) sheet.getRow(rownumber).getCell(colnumber).getNumericCellValue();
				
				data=Integer.toString(numvalue);
				
				break;
			}
			
			
			
			
			
		
		
		
		return data;
		
		
	}
	
	
	
	
	
	
	
	
	public static void loadTestData()
	{
		
		System.out.println("Loading Excel");
		// step 1: Get the filepath
		
		
				String filepath=System.getProperty("user.dir")+"\\Testdata\\Testdata.xlsx";
						
				
				File file=new File(filepath);
				
				
				// User Reader to read the excel file
				
				FileInputStream fis;
				try {
					fis = new FileInputStream(file);
					

					// Create WorkBook Instance
					
					workbook=new XSSFWorkbook(fis);
					
				} catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				
				System.out.println("Loading Excel is completed");
				
	}
	
	
	
	
	

}
