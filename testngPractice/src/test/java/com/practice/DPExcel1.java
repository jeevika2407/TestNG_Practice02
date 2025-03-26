package com.practice;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;
public class DPExcel1 {
  @DataProvider(name="excelData",parallel=true)
  public Object[][] excelDP() throws IOException{
	  
	  String loc=System.getProperty("user.dir")+"/src/test/resources/testdata1.xlsx";
	  Object[][] obj=getData(loc,"Sheet1");
	  return obj;
  }
  public String[][] getData(String file,String sheet) throws IOException{
	  String[][] data=null;
	  try {
	  FileInputStream fis=new FileInputStream(file);
	  XSSFWorkbook wrkBk=new XSSFWorkbook(fis);
	  XSSFSheet sht=wrkBk.getSheet(sheet);
	  XSSFRow row=sht.getRow(0);
	  int r=sht.getPhysicalNumberOfRows();
	  int c=row.getLastCellNum();
	  Cell cell;
	  data=new String[r][c];
	  for(int i=0;i<r;i++) {
		  for(int j=0;j<c;j++){
			  row=sht.getRow(i);
			  cell=row.getCell(j);
			  data[i][j]=cell.getStringCellValue();
		  }
	  }
	  }
	  catch(Exception e){
		  System.out.println(e.getMessage());
	  }
	  return data;
  }
}