package Demo;

import java.io.FileInputStream;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Demo11 {

	public static void main(String[] args) throws Exception {
		
		String Sheetname="Sheet1";
		
		FileInputStream file = new FileInputStream("C:\\Users\\Tulasikumar\\eclipse-workspace\\44444444\\src\\test\\java\\Demo\\New folder\\Folder.xlsx");
		
		XSSFWorkbook book =new XSSFWorkbook(file);
		XSSFSheet Sheet = book.getSheet(Sheetname);
		
		for(int i=0;i<Sheet.getLastRowNum()+1;i++) {
		
		for(int j=0;j<Sheet.getRow(i).getLastCellNum();j++){
			
			String data = Sheet.getRow(i).getCell(j).toString();
			
			System.out.print(" "+data);
		}
System.out.println();
	}

}
}
