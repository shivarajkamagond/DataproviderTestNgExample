package testingpurpose;

import org.testng.annotations.Test;
import org.testng.annotations.DataProvider;

public class Exceltodataprovider {
	ExcelOperations eat=null;
	String xlfilepath="C:\\Users\\acer pc\\Desktop\\testingdataprovider.xlsx";
	int index;
	
	@Test(dataProvider = "userData")
	public void fillform(String name,String password,String results){
		System.out.println("name :"+name);
		System.out.println("password :"+password);
		System.out.println("results :"+results);
		
		System.out.println("*********************");

	}


	
	
       
	@DataProvider(name="userData")
	 public Object[][]userFormData() throws Exception{
		index=0;
		Object[][] data=testData(xlfilepath,index);
		return data; 
	}   
	
	@DataProvider(name="Data")
	 public Object[][] Data() throws Exception{
		index=1;
		Object[][] data=testData(xlfilepath,index);
		return data; 
	}
	
	
	public Object[][] testData(String xlfilepath,int index) throws Exception{
		Object excelData[][];
		eat=new ExcelOperations(xlfilepath);
          
		int rows= eat.GetTotalRowCount(index);
		int columns=eat.GetTotalColumnCount(index);
		
		excelData=new Object[rows-1][columns];
		for(int i=1;i<rows;i++){
			for(int j=0;j<columns;j++){
				
				excelData[i-1][j]=ExcelOperations.getcelldata( index,i,j);
//			System.out.println(ExcelOperations.GetCelldata( index,i,j));
			}
		}
			return excelData;
	}


}
