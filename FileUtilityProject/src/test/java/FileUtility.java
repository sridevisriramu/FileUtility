import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class FileUtility {

	public static void main(String[] args) throws Exception {

		FileUtility fu = new FileUtility();
		fu.testExcelRemoveRow();
		fu.testExcelOperationsAddRow();
	}

	public void testExcelOperationsAddRow() throws IOException {
		XSSFWorkbook workbook = null;

		String path = "C:/test/test.xlsx";
		File file = new File(path).getAbsoluteFile();
		
		System.out.println("file = "+ file);

		FileInputStream fileInputStream = new FileInputStream(file);
		
		System.out.println("fileInputStream = "+ fileInputStream);
		workbook = new XSSFWorkbook(fileInputStream);

		XSSFSheet worksheet =  workbook.getSheetAt(1);
		//int cellNumber = worksheet.getRow(0).getLastCellNum();
		Row row = worksheet.createRow(1);
		
		for(int i = 0 ; i < 10; i++)
		{
			Cell cell = row.createCell(i, 1);
			cell.setCellValue(i);
		}

		FileOutputStream hOutPut = new FileOutputStream(file);
		workbook.write(hOutPut);
		
		workbook.close();
	}

	
	
	public void testExcelRemoveRow() throws Exception
	{
		XSSFWorkbook workbook = null;
		String path = "C:/test/test.xlsx";
		try
		{
		File file = new File(path).getAbsoluteFile();
		FileInputStream inputStream = new FileInputStream(file);
		workbook = new XSSFWorkbook(inputStream);
		XSSFSheet sheet = workbook.getSheetAt(0);
		int count = sheet.getPhysicalNumberOfRows() - 1;
		for (int i = 1; i <= count; i++) {
			Row row = sheet.getRow(i);
			sheet.removeRow(row);
		}

		FileOutputStream out = new FileOutputStream(file);
		workbook.write(out);
		System.out.println("test completed");
		} catch (Exception e) {
			//log.error("Error came while reading from excel sheet, Failed");
			throw new Exception(e);
		} finally {
			workbook.close();
			System.out.println("test completed");
		}
	}
}
