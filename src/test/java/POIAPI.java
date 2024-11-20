import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class POIAPI {

	public static void main(String[] args) throws IOException {
		
		File excelFile=new File(System.getProperty("user.dir")+"\\src\\test\\resources\\Employees.xlsx");
		FileInputStream fis=new FileInputStream(excelFile);
		
		XSSFWorkbook workbook=new XSSFWorkbook(fis);
		XSSFSheet sheet=workbook.getSheet("Basic details");  //sheet name
		
		int rows=sheet.getLastRowNum();   //no of rows
		System.out.println(rows);
		
		int cells=sheet.getRow(0).getLastCellNum();
		System.out.println(cells);     //no of cells
		
		for(int r=0;r<rows;r++)
		{
			XSSFRow row=sheet.getRow(r);
			
			for(int c=0;c<cells;c++)
			{
				XSSFCell cell = row.getCell(c);
				CellType celltype = cell.getCellType();
				
				switch(celltype)
				{
				  case NUMERIC:
					  System.out.print(cell.getNumericCellValue()+" || ");
				   break;
				   
				  case STRING:
					  System.out.print(cell.getStringCellValue()+" || ");
					  break;
					  
				  case BOOLEAN:
					  System.out.print(cell.getBooleanCellValue()+" || ");
					  break;
					  
			      default:
			    	  System.out.print("Data is not matching with given cases");
				   
				 }
				
			}
			
			System.out.println();
		}
		workbook.close();
	}
	
}
		