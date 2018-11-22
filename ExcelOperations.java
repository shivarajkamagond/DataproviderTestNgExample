package testingpurpose;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class ExcelOperations {
	public FileInputStream fis=null;
	public static XSSFWorkbook  workbook=null;
	public static  XSSFSheet  sheet=null;
	public static XSSFRow row=null;
	public static XSSFCell cell=null;
	
	public ExcelOperations(String xlFilePath) throws Exception {
		fis=new FileInputStream(xlFilePath);
		workbook=new XSSFWorkbook(fis);
		fis.close();
	}
	@Test
	public static String GetCelldata(int index,int rownum,int columnnum) throws IOException
	{
		String rowdata="";
		sheet=workbook.getSheetAt(index);
		
		rowdata=sheet.getRow(rownum).getCell(columnnum).toString();
		//System.out.println(rowdata);
		return rowdata;
	}

	public static String getcelldata(int index,int rownum,int columnnum)throws IOException
	{
		try {
			sheet=workbook.getSheetAt(index);
			row=sheet.getRow(rownum);
			cell=row.getCell(columnnum);
			
			if(cell.getCellTypeEnum()== CellType.STRING)
				return cell.getStringCellValue();
			else if(cell.getCellTypeEnum()== CellType.NUMERIC ||cell.getCellTypeEnum()== CellType.FORMULA) {
				String cellValue=String.valueOf(cell.getNumericCellValue());
				if(HSSFDateUtil.isCellDateFormatted(cell) ) {
					DateFormat dt= new SimpleDateFormat("dd/MM/yy");
					Date date= cell.getDateCellValue();
					cellValue=dt.format(date);
					
				}
				return cellValue;
			}
			else if(cell.getCellTypeEnum()== CellType.BLANK )
				return "";
			else
				return String.valueOf(cell.getBooleanCellValue());
			
		}
	
		catch (Exception e) {
			e.printStackTrace();
			return "no match found";
		}
		
	}

	
//   public int getRowCount(int index) {
//	   sheet=workbook.getSheetAt(index);
//	   	   int rowCount=0;
//	   rowCount= sheet.getLastRowNum()+1;
//	return (rowCount);
//	   
//   }
   public int GetTotalColumnCount(int index) throws IOException
	{
		//int rownum=0;
		
		
	sheet=workbook.getSheetAt(index);
	    row=sheet.getRow(0);
		int ColumnCount=row.getLastCellNum();
	//	System.out.println(ColumnCount);
		
		return ColumnCount;
	}
   
//   public int getColumnCount(int index) {
//	   sheet=workbook.getSheetAt(index);
//	   row=sheet.getRow(0);
//	   int col=row.getLastCellNum();
//	   
//	return col;
//	}
   public int GetTotalRowCount(int index) throws IOException
	{
		
     sheet=workbook.getSheetAt(index);
		
		int TotalRow=sheet.getLastRowNum()+1;
		return TotalRow;
	}
}
