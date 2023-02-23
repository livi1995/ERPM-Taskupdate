package teams;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excel {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		String downloadFilepath = "Z:\\MECHANICAL\\TEAM (USERS)\\LIVINGSTON DURAI\\Tasks automation\\Project folders\\Default path";
		String destinationDir = "Z:\\MECHANICAL\\TEAM (USERS)\\LIVINGSTON DURAI\\Tasks automation\\Project folders\\";
		String taskid = "T123";
		destinationDir=destinationDir+taskid;
		moveFilesToNewFolder(downloadFilepath, destinationDir);
		String excelPath="Z:\\MECHANICAL\\TEAM (USERS)\\LIVINGSTON DURAI\\Tasks automation\\Task sheet.xlsx";
		String stringToCheck="T6472";
		checkAndAddString(excelPath, stringToCheck);

	}
	
	public static void checkAndAddString(String excelPath, String stringToCheck) throws IOException {
        // Open the Excel file
        FileInputStream inputStream = new FileInputStream(new File(excelPath));
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0); // assuming you want to access first sheet in excel

        boolean stringExist = false;
        // Iterate over the rows in the sheet
        for (Row row : sheet) {
            Cell cell = row.getCell(1); // get column B
            if (cell != null) {
                String cellValue = cell.getStringCellValue();
                // Check if the string exists in column B
                if (cellValue.equals(stringToCheck)) {
                    stringExist = true;
                    break;
                }
            }
        }

        // if string doesn't exist in column B, add it to a new row
        if (!stringExist) {
            int lastRow = sheet.getLastRowNum();
            Row newRow = sheet.createRow(lastRow + 1);
            Cell newCell = newRow.createCell(1);
            newCell.setCellValue(stringToCheck);
            
            Cell colCell = newRow.createCell(2);
            colCell.setCellValue("Title");
            colCell = newRow.createCell(3);
            colCell.setCellValue("Descr");
            colCell = newRow.createCell(4);
            colCell.setCellValue("Due date");
            colCell = newRow.createCell(5);
            colCell.setCellValue("Priori");
        }

        // Close the Excel file
        inputStream.close();

        // Write the changes to the Excel file
        FileOutputStream outputStream = new FileOutputStream(excelPath);
        workbook.write(outputStream);
        outputStream.close();
        workbook.close();
    }
	
	 public static void moveFilesToNewFolder(String sourceDir, String destinationDir) {
	        // Check if the destination directory exists
	        File destinationFolder = new File(destinationDir);
	        if (!destinationFolder.exists()) {
	            // Create the destination directory
	            destinationFolder.mkdir();
	        }

	        // Get a list of all files in the source directory
	        File[] files = new File(sourceDir).listFiles();

	        // Iterate over the files and move them to the destination directory
	        for (File file : files) {
	            File newFile = new File(destinationDir + "/" + file.getName());

	            // Move the file to the new location
	            boolean success = file.renameTo(newFile);

	            if (success) {
	                System.out.println("Successfully moved file " + file.getName() + " to " + destinationDir);
	            } else {
	                System.out.println("Failed to move file " + file.getName() + " to " + destinationDir);
	                
	            }}}

}
