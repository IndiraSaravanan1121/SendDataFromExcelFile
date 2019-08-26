package ReadExcelData.ReadExcelDataSelenium;

import java.io.FileInputStream;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

/*
 * Purpose:Take the value from excel file and passing the value to website.
 * Apache poi Supports to read value from excel file.
 */
public class SendValueFromExcel {
	
	public static void main(String[] args) throws Exception {
		
		// configure the driver
		System.setProperty("webdriver.chrome.driver",
				"C:\\Users\\indira.saravanan\\eclipse-workspace\\Selenium\\driver\\chromedriver_win32\\chromedriver.exe");
		
		// create new object for chrome driver
		WebDriver driver = new ChromeDriver();
		
		//Mention URL
		driver.get("https://www.orangehrm.com/");

		driver.findElement(By.xpath("//*[@id=\"header-navbar\"]/ul[2]/li[1]/a")).click();

		//create an object of FileInputStream class to read excel file.
		FileInputStream file = new FileInputStream("C:\\Users\\indira.saravanan\\Documents\\book.xlsx");
		
		//To access workbook,create object for that.
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		
		//To access particular sheet
		XSSFSheet sheet = workbook.getSheet("sheet1");
		int noOfRows = sheet.getLastRowNum();
		System.out.println("No.of records in excel sheet:" + noOfRows);//no of rows in excel file 
		
		//create loop to read all the row in excel file.
		for (int row = 1; row <= noOfRows; row++) {
			XSSFRow currentRow = sheet.getRow(row);

			String First_name = currentRow.getCell(0).getStringCellValue();
			String Second_name = currentRow.getCell(1).getStringCellValue();
			String Company_name = currentRow.getCell(2).getStringCellValue();
			String No_of_employee = currentRow.getCell(3).getStringCellValue();
			String Phone = currentRow.getCell(4).getStringCellValue();
			String JobTitle = currentRow.getCell(5).getStringCellValue();
			String Email = currentRow.getCell(6).getStringCellValue();
			String Country = currentRow.getCell(7).getStringCellValue();
			String Comment = currentRow.getCell(8).getStringCellValue();

			
			//Sending value to website.
			driver.findElement(By.name("firstname")).sendKeys(First_name);
			driver.findElement(By.name("lastname")).sendKeys(Second_name);
			driver.findElement(By.name("company")).sendKeys(Company_name);

			Select dropCountry = new Select(driver.findElement(By.name("numemployees")));
			dropCountry.selectByVisibleText(No_of_employee);

			driver.findElement(By.name("phone")).sendKeys(Phone);
			driver.findElement(By.name("jobtitle")).sendKeys(JobTitle);
			driver.findElement(By.name("email")).sendKeys(Email);

			Select dropCountry1 = new Select(driver.findElement(By.name("country")));
			dropCountry1.selectByVisibleText(Country);

			driver.findElement(By.name("message")).sendKeys(Comment);

			//To close driver.
			driver.close();
		}

	}

}
