package poi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

// This program copy/paste an Excel .xlsm with macros and remove one sheet 
// Important have imported Apache POI, and commons-collections in the library path

public class RemoveSheetXlsm {

	public static void main(String[] args) {

		
		InputStream in = null;
		XSSFWorkbook workbook = null;

		try {
			in = new FileInputStream("Proyecto1.xlsm");
			workbook = new XSSFWorkbook(in);

		} catch (IOException e) {
			System.err.println("Error reading XLSM Input File");
			e.printStackTrace(System.err);
		}
		//more code here


		workbook.removeSheetAt(2); //in this example, remove sheet nº3 (init = 0)

		//output is an .xlsm file also and contains the same Excel macros without removed sheet
		FileOutputStream outputfile;
		try {
			outputfile = new FileOutputStream("OutputFile.xlsm");
			workbook.write(outputfile);
			outputfile.flush();
			outputfile.close();
			workbook.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace(System.err);
		} catch (IOException e) {
			e.printStackTrace(System.err);
		} catch (Exception e) {
			e.printStackTrace(System.err);
		}

	}

}
