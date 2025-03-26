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

/*
public Object[][] getData(String file, String sheet) throws IOException {
    Object[][] data = null;

    try (FileInputStream fis = new FileInputStream(file);
         XSSFWorkbook wrkBk = new XSSFWorkbook(fis)) {

        XSSFSheet sht = wrkBk.getSheet(sheet);
        int rowCount = sht.getPhysicalNumberOfRows();
        int colCount = sht.getRow(0).getLastCellNum();

        data = new Object[rowCount][colCount];

        for (int i = 0; i < rowCount; i++) {
            XSSFRow row = sht.getRow(i);
            for (int j = 0; j < colCount; j++) {
                XSSFCell cell = row.getCell(j);

                switch (cell.getCellType()) {
                    case NUMERIC:
                        data[i][j] = (int) cell.getNumericCellValue(); // Convert to int
                        break;
                    case STRING:
                        data[i][j] = cell.getStringCellValue();
                        break;
                    case BOOLEAN:
                        data[i][j] = cell.getBooleanCellValue();
                        break;
                    case FORMULA:
                        data[i][j] = cell.getCellFormula();
                        break;
                    default:
                        data[i][j] = ""; // Handle empty cells
                        break;
                }
            }
        }
    } catch (Exception e) {
        System.out.println("Error reading Excel: " + e.getMessage());
    }

    return data;
    */
}
}