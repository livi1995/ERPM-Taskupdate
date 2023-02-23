package erpm_Task_Update;

import java.awt.HeadlessException;
import java.awt.Toolkit;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.io.File;
import java.io.IOException;
import java.time.Duration;
import java.util.HashMap;
import java.util.List;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class Task_update {
	public static WebDriver driver;
	public static String downloadFilepath = "Z:\\MECHANICAL\\TEAM (USERS)\\LIVINGSTON DURAI\\Tasks automation\\Project folders\\Default path";
	public static String excelPath="Z:\\MECHANICAL\\TEAM (USERS)\\LIVINGSTON DURAI\\Tasks automation\\Task sheet.xlsx";
	public static void main(String[] args) throws HeadlessException, UnsupportedFlavorException, IOException {
		// TODO Auto-generated method stub
		try {
		System.setProperty("webdriver.chrome.driver", "D:\\Livingston\\Selenium\\chromedriver_win32\\chromedriver.exe");
		
		String downloadFilepath = "Z:\\MECHANICAL\\TEAM (USERS)\\LIVINGSTON DURAI\\Tasks automation\\Project folders\\Default path";
		HashMap<String, Object> chromePrefs = new HashMap<String, Object>();
		chromePrefs.put("profile.default_content_settings.popups", 0);
		chromePrefs.put("download.default_directory", downloadFilepath);
		chromePrefs.put("safebrowsing.enabled", "true");
		chromePrefs.put("download.prompt_for_download", "False");
		chromePrefs.put("download.directory_upgrade", "True");
		chromePrefs.put("safebrowsing.enabled", "False");
		
		ChromeOptions options = new ChromeOptions();
		//options.addArguments("user-data-dir=C:\\Users\\20309017\\AppData\\Local\\Google\\Chrome\\User Data\\Default");
		options.addArguments("--start-maximized");
		options.setExperimentalOption("prefs", chromePrefs);
		driver = new ChromeDriver(options);
		
		//WebDriver driver = new ChromeDriver();
		
		     			
		driver.get("https://erpm.ltmindia.com/homepage");
		//driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(20));
		//Thread.sleep(10000);
		WebElement eL =driver.findElement(By.name("UserName"));
		eL.sendKeys("asreng.ltrpm@larsentoubro.com");
		eL.sendKeys(Keys.RETURN);
		//Thread.sleep(3000);
		eL =driver.findElement(By.id(":r1:"));
		eL.sendKeys("Asrpu@123");
		eL.sendKeys(Keys.RETURN);
		
		//Thread.sleep(10000);
		
		driver.findElement(By.cssSelector("img[alt='user.displayName'")).click();
		Thread.sleep(1000);
		driver.findElement(By.xpath("//*[contains(text(),'Boards')]")).click();
		Thread.sleep(5000);
		
        
		setStatus(excelPath);
		//setStatus(excelPath, 1);
		execute(driver, "NPD");
		execute(driver, "spare");
		System.out.println("File updated");
		driver.quit();
		
		}catch (InterruptedException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
		
	public static void execute(WebDriver d, String s) throws InterruptedException, HeadlessException, UnsupportedFlavorException, IOException {
		//String str;
		
		if (s=="NPD") {
		driver.findElement(By.xpath("//h4[text()='Enquiry - ASR Eqp/Aux']")).click();
		}else {
		driver.findElement(By.xpath("//h4[text()='ASR Spares']")).click();
		}
		Thread.sleep(10000);
		//get all task element
		
		List<WebElement> tasks;
		if (s=="NPD") {
			 tasks = driver.findElements(By.xpath("//div[@data-rbd-droppable-id='73']//div[@style='height: fit-content;']"));
		}else {
			 tasks = driver.findElements(By.xpath("//div[@data-rbd-droppable-id='79']//div[@style='height: fit-content;']"));
		}
		
		
		Thread.sleep(5000);
		int i = 1;
		//iterate for each element
		for (WebElement task : tasks)
		{
			
		task.click();
		Thread.sleep(3000);
		
		//get task details
		String title=( driver.findElement(By.xpath("//input[@name='task.taskName']")).getAttribute("value"));
		String taskid=( driver.findElement(By.xpath("(//h4[contains(text(),T)])[3]")).getText());
		Thread.sleep(1000);
		driver.findElement(By.xpath("//*[contains(text(),'History')]")).click();
		
		//List<WebElement> count = driver.findElements(By.xpath("//*[@class=\"MuiTypography-root MuiTypography-caption css-mzy3wy\"]"));
		
		String startdate =driver.findElement(By.xpath("//*[text()='created the task']/following-sibling::span")).getText();
		String duedate=( driver.findElement(By.xpath("//input[@class='MuiOutlinedInput-input MuiInputBase-input MuiInputBase-inputSizeSmall MuiInputBase-inputAdornedEnd css-b52kj1']")).getAttribute("value"));
		String proirity=( driver.findElement(By.xpath("(//input[contains(id,mui-406)])[9]")).getAttribute("value"));
		String webid=( driver.findElement(By.xpath("//*[contains(@id,\"buttonfile\")]")).getAttribute("id"));
		webid= webid.substring(10);
		webid= "https://erpm.ltmindia.com/task/detail/"+ webid;
		Thread.sleep(1000);
		driver.switchTo().frame(0);
		WebElement textBox = driver.findElement(By.xpath("//body[@id='tinymce']"));
		textBox.sendKeys(Keys.CONTROL + "a");
		Thread.sleep(1000);

        textBox.sendKeys(Keys.CONTROL + "c");
        Thread.sleep(1000);
        driver.switchTo().defaultContent();
        Thread.sleep(1000);
        String data = (String) Toolkit.getDefaultToolkit().getSystemClipboard().getData(DataFlavor.stringFlavor); 
        //System.out.print(data);
        Thread.sleep(1000);
        String stringToCheck = taskid;
        
		//write data to excel
        boolean stringExist;
        if (s=="NPD") {
        	stringExist = checkAndAddString(excelPath, stringToCheck,title,data,startdate,duedate,proirity,webid,0);
		}else {
			stringExist = checkAndAddString(excelPath, stringToCheck,title,data,startdate,duedate,proirity,webid,1);
		}
        
        //clear clipboard data
        
        //download all attachment
        List<WebElement> attachments = driver.findElements(By.xpath("//*[@class=\"MuiList-root MuiList-dense css-1uzmcsd\"]//div[@role=\"button\"]//span[@class=\"MuiTypography-root MuiTypography-body2 MuiListItemText-primary css-16nsi8u\"]"));
		Thread.sleep(3000);
		
		System.out.println("Task - "+ stringToCheck+" exist -->"+stringExist);
		if(stringExist==false)
		{
		for (WebElement attachment : attachments)
		{
			
			String a=attachment.getText();	
			if (a.toLowerCase().endsWith(".jpg")) {
	            System.out.print(a);
	            attachment.click();
	            takeScreenshot(driver, downloadFilepath+"\\"+ attachment.getText());
	            Thread.sleep(5000);
	            Actions act = new Actions(driver);
	            act.sendKeys(Keys.chord(Keys.ESCAPE)).perform();
	            //attachment.sendKeys(Keys.ESCAPE);
	            //driver.findElement(By.xpath("//div[@class='MuiBox-root css-1a3f7hq']")).sendKeys(Keys.ESCAPE);
	        }
			else {
				attachment.click();
			}
			Thread.sleep(2000);
		}
		Thread.sleep(3000);
		//moving downloaded files
		String destinationDir = "Z:\\MECHANICAL\\TEAM (USERS)\\LIVINGSTON DURAI\\Tasks automation\\Project folders\\";
		destinationDir=destinationDir+taskid;
		moveFilesToNewFolder(downloadFilepath, destinationDir);
		}
        //click close button
		if (s=="NPD") {
        driver.findElement(By.xpath("(//button[@type=\"button\"])[19]")).click();
		}else {
			driver.findElement(By.xpath("//*[@class='MuiButtonBase-root MuiIconButton-root MuiIconButton-edgeEnd MuiIconButton-sizeLarge css-9wr0ai']")).click();
			//MuiButtonBase-root MuiIconButton-root MuiIconButton-edgeEnd MuiIconButton-sizeLarge css-9wr0ai
		}
		i=i+1;
		Thread.sleep(3000);
		}
		driver.findElement(By.xpath("//*[@d='M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z']")).click();
		
		} 
		

	
	 public static WebElement waitForElement(WebDriver driver, By locator, Duration timeout) 
	 {
	        WebDriverWait wait = new WebDriverWait(driver, timeout);
	        return wait.until(ExpectedConditions.visibilityOfElementLocated(locator));
	    }
	 
	 public static void setStatus(String excelPath) throws IOException {
	        // Open the Excel file
	        FileInputStream inputStream = new FileInputStream(new File(excelPath));
	        Workbook workbook = new XSSFWorkbook(inputStream);
	        for(int i=0;i<2;i++)
	        {
	        Sheet sheet = workbook.getSheetAt(i); // assuming you want to access first sheet in excel

	        // Iterate over the rows in the sheet
	        for (Row row : sheet) {
	            Cell status = row.getCell(8);
	            status.setCellValue("COMPLETED");
	                }
	            }
	        FileOutputStream outputStream = new FileOutputStream(excelPath);
	        workbook.write(outputStream);
	        outputStream.close();
	        workbook.close();
	 }
	        


	 public static boolean checkAndAddString(String excelPath, String stringToCheck, String title, String Descr,String startdate, String DueDate,String Priority, String webid, int sheetNo) throws IOException {
	        // Open the Excel file
	        FileInputStream inputStream = new FileInputStream(new File(excelPath));
	        Workbook workbook = new XSSFWorkbook(inputStream);
	        Sheet sheet = workbook.getSheetAt(sheetNo); // assuming you want to access first sheet in excel

	        boolean stringExist = false;
	        // Iterate over the rows in the sheet
	        for (Row row : sheet) {
	            Cell cell = row.getCell(1); // get column B
	            
	            
	            if (cell != null) {
	                String cellValue = cell.getStringCellValue();
	                // Check if the string exists in column B
	                if (cellValue.equals(stringToCheck)) {
	                    stringExist = true;
	                    Cell status = row.getCell(8);
	                    status.setCellValue("WIP");
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
	            colCell.setCellValue(title);
	            colCell = newRow.createCell(3);
	            colCell.setCellValue(Descr);
	            colCell = newRow.createCell(4);
	            colCell.setCellValue(startdate);
	            colCell = newRow.createCell(5);
	            colCell.setCellValue(DueDate);
	            colCell = newRow.createCell(6);
	            colCell.setCellValue(Priority);
	            colCell = newRow.createCell(7);
	            colCell.setCellValue(webid);
	            colCell = newRow.createCell(8);
	            colCell.setCellValue("WIP");
	        }

	        // Close the Excel file
	        inputStream.close();

	        // Write the changes to the Excel file
	        FileOutputStream outputStream = new FileOutputStream(excelPath);
	        workbook.write(outputStream);
	        outputStream.close();
	        workbook.close();
	        return stringExist;
	    }
		
		 public static void moveFilesToNewFolder(String sourceDir, String destinationDir ) {
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
		                file.delete();
		                
		            }}}
		 public static void takeScreenshot(WebDriver driver, String fileName) throws IOException {
		        File screenshot = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
		        FileUtils.copyFile(screenshot, new File(fileName));
		    }

	        
	     }

	    
	         
	     
	 


