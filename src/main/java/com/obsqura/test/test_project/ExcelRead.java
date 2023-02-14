package com.obsqura.test.test_project;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {
public static void main(String args[])
	{
List<Double> list =new ArrayList<Double>();
File file=new File("C:\\Users\\angit\\OneDrive\\Desktop\\Obsqura Training\\Assignment and Notes\\Test Excel.xlsx");
FileInputStream fis = null;
try 
{
	fis = new FileInputStream(file);
} 
catch (FileNotFoundException e) 
{
	// TODO Auto-generated catch block
	e.printStackTrace();
}
XSSFWorkbook wb = null;
try 
{
	wb = new XSSFWorkbook(fis);
}
catch (IOException e) 
{
	// TODO Auto-generated catch block
	e.printStackTrace();
}
XSSFSheet sheet=wb.getSheetAt(0);
Iterator<Row> itr=sheet.iterator();
while(itr.hasNext())
{
	Row row=itr.next();
	Iterator<Cell> cellIterator=row.cellIterator();
	while(cellIterator.hasNext())
	{
		Cell cell=cellIterator.next();
		switch(cell.getCellType())
		{
		case STRING:
		System.out.println(cell.getStringCellValue()+"\t\t\t");
		break;
		case NUMERIC:
		System.out.println(cell.getNumericCellValue()+"\t\t\t");
		list.add(cell.getNumericCellValue());
		default:
		}
		
	}	
	}
double principal=(Double)list.get(0);
double noofyears=(Double)list.get(1);
double interestrate=(Double)list.get(2);
double interestamt=principal*noofyears*interestrate;
System.out.println("Interest is Rs."+interestamt);
}
}
