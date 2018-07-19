package poi;


import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Workbook;

// This program create an XLS file with 2 named sheets, name a cell range and put data on it.

public class CreateXLS {

	public static void main(String[] args) throws IOException {
				
			HSSFWorkbook workbook = new HSSFWorkbook(); //make workbook
		    

	        HSSFSheet sheet = workbook.createSheet(); //first sheet
	        HSSFSheet sheet1 = workbook.createSheet(); //second sheet

	        workbook.setSheetName(0, "Name1"); //name first sheet
	        workbook.setSheetName(1, "Name2"); //name second sheet

		  
	        sheet.createRow(0).createCell((short) 0).setCellValue("hola");

		    // Create named range for a single cell using areareference
		    Name namedCell = workbook.createName();
		    namedCell.setNameName("celda"); //name of the cell range
		    String reference = "Biologic!$A$1:$A$1"; // area reference
		    namedCell.setRefersToFormula(reference);

		    // Output file
	        FileOutputStream file = new FileOutputStream("Test.xls");

	        workbook.write(file);

	        file.close();


	}

}
