import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

import java.awt.*;
import java.awt.event.KeyEvent;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

public class ReactJS_Demo_With_Gradle {

	WebDriver driver;

	@BeforeMethod
	public void setup() {

		System.setProperty("webdriver.chrome.driver", "C:/Program Files/Google/Chrome/Application/chromedriver.exe");
		driver = new ChromeDriver();
//	  System.setProperty("webdriver.gecko.driver", "D:\\geckodriver.exe");

//	  driver = new FirefoxDriver();


	}

	@Test(priority = 0)
	public void executeScript() throws InterruptedException, AWTException, IOException {

		int rowCounter = 0;
		//declare file name to be create
		String filename = "D:\\ReactJS_Demo_Web_Application.xls";
		//creating an instance of HSSFWorkbook class
		HSSFWorkbook workbook = new HSSFWorkbook();

		HSSFSheet sheet = workbook.createSheet("Test Cases");

		try {


			HSSFRow rowhead = sheet.createRow(rowCounter);
			//creating cell by using the createCell() method and setting the values to the cell by using the setCellValue() method
			rowhead.createCell(0).setCellValue("Test Case No.");
			rowhead.createCell(1).setCellValue("Test Case Description.");
			rowhead.createCell(2).setCellValue("Result");
			rowCounter++;

			//Launching the Site.
			// driver.get("http://e6d3c54a-fe82-4282-b680-a2e90e8e3776.cloudapp.net/");
			//creating the 1st row
			HSSFRow row = sheet.createRow(rowCounter);
			//inserting data in the first row
			System.out.println("1st Test Case");
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to check whether the application opens with the given URL and to verfiy the title");
			//String result = "Pass";


			System.out.println("Test case to check whether the application opens with the given URL and to verfiy the title");

			driver.get("http://71c8ac46-0c00-439b-b756-e06e05c47f24.cloudapp.net/");

			if(driver.getTitle().equalsIgnoreCase("CloudEQ App"))
			{
				String result = "Pass";
				System.out.println("This test case is passed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
				driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			}

			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
			}


			//Maximize window
			driver.manage().window().maximize();

			//Set the Script Timeout to 20 seconds
			driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);


			System.out.println("2nd Test Case");
			row = sheet.createRow(rowCounter);
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to check whether the Home Tab is displayed and clickable");

			//Navigating to each page
			if(driver.findElement(By.xpath("//a[contains(text(),'Home')]")).isDisplayed())
			{
				System.out.println("Test case to check whether the Home Tab is displayed and clickable");

				WebElement home_button = driver.findElement(By.xpath("//a[contains(text(),'Home')]"));
				home_button.click();
				String result = "Pass";
				System.out.println("This test case is passed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
				driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			}
			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
			}

			System.out.println("3rd Test Case");
			row = sheet.createRow(rowCounter);
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to check whether the Department Tab is displayed and clickable");

			if(driver.findElement(By.xpath("//a[contains(text(),'Department')]")).isDisplayed())
			{
				System.out.println("Test case to check whether the Department Tab is displayed and clickable");
				WebElement department_button = driver.findElement(By.xpath("//a[contains(text(),'Department')]"));
				department_button.click();
				String result = "Pass";
				System.out.println("This test case is passed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
				driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			}
			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
			}

			System.out.println("4th Test Case");
			row = sheet.createRow(rowCounter);
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to check whether the Employee Tab is displayed and clickable");

			if(driver.findElement(By.xpath("//a[contains(text(),'Employee')]")).isDisplayed())
			{
				System.out.println("Test case to check whether the Employee Tab is displayed and clickable");
				WebElement employee_button = driver.findElement(By.xpath("//a[contains(text(),'Employee')]"));
				employee_button.click();
				String result = "Pass";
				System.out.println("This test case is passed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
				driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			}
			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
			}

			// Testing each page redirects to the right page
			// Now Testing that after clicking on Home button, does it redirects to the home page

			System.out.println("4th Test Case");
			row = sheet.createRow(rowCounter);
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to check whether the Home page is displayed");

			System.out.println("Test case to check whether the Home page is displayed");
			WebElement home_button2 = driver.findElement(By.xpath("//a[contains(text(),'Home')]"));
			home_button2.click();
			driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			if(driver.findElement(By.xpath("//h3[contains(text(),'Welcome to CloudEQ Management Portal. This is the demo for CI CD.')]")).getText().contains("CloudEQ"))
			{
				String result = "Pass";
				System.out.println("This test case is passed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
				driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			}
			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
			}

			// Now Testing that after clicking on Department button, does it redirects to the Department page

			System.out.println("5th Test Case");
			row = sheet.createRow(rowCounter);
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to check whether the Department page is displayed");

			System.out.println("Test case to check whether the Department page is displayed");
			WebElement department_button = driver.findElement(By.xpath("//a[contains(text(),'Department')]"));
			department_button.click();

			driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			if(driver.findElement(By.xpath(".//th[contains(text(),'DepartmentName')]")).getText().contains("DepartmentName"))
			{
				String result = "Pass";
				System.out.println("This test case is passed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
				driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			}
			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
			}

			// Now Testing that after clicking on Employee button, does it redirects to the Employee page

			System.out.println("6th Test Case");
			row = sheet.createRow(rowCounter);
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to check whether the Employee page is displayed");

			System.out.println("Test case to check whether the Employee page is displayed");
			WebElement employee_button = driver.findElement(By.xpath("//a[contains(text(),'Employee')]"));
			employee_button.click();
			driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			if(driver.findElement(By.xpath(".//th[contains(text(),'EmployeeName')]")).getText().contains("EmployeeName"))
			{
				String result = "Pass";
				System.out.println("This test case is passed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
				driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			}
			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
			}
//  }

			// @Test(priority = 1)
			// public void testDepartmentPage() throws InterruptedException, AWTException {

			driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);

			JavascriptExecutor js = (JavascriptExecutor)driver;

			// Clicking on Department tab

			System.out.println("7th Test Case");
			row = sheet.createRow(rowCounter);
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to check whether the Department Tab is clickable");

			System.out.println("Test case to check whether the Department Tab is clickable");
			if(driver.findElement(By.xpath("//a[contains(text(),'Department')]")).isDisplayed())
			{
				WebElement department_button2 = driver.findElement(By.xpath("//a[contains(text(),'Department')]"));
				department_button2.click();
				String result = "Pass";
				System.out.println("This test case is passed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
				driver.manage().timeouts().implicitlyWait(5000, TimeUnit.SECONDS);
				js.executeScript("window.scrollTo(0, document.body.scrollHeight)");
				//driver.findElement(By.cssSelector("html")).sendKeys(Keys.CONTROL, Keys.END);
			}
			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
			}

			// Clicking on Add department button
			System.out.println("8th Test Case");
			row = sheet.createRow(rowCounter);
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to check whether the Add Department button is clickable");

			System.out.println("Test case to check whether the Add Department button is clickable");
			if(driver.findElement(By.xpath("//button[contains(text(),'Add Department')]")).isDisplayed())
			{
				WebElement add_department_button = driver.findElement(By.xpath("//button[contains(text(),'Add Department')]"));
				add_department_button.click();
				String result = "Pass";
				System.out.println("This test case is passed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
				//driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
				Thread.sleep(100L);
			}
			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
			}



			// Adding New department

			System.out.println("9th Test Case");
			row = sheet.createRow(rowCounter);
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to check whether the Department name is displayed in the Add department dialog box");

			System.out.println("Test case to check whether the Department name is displayed in the Add department dialog box");
			if(driver.findElement(By.xpath("//input[@id='DepartmentName']")).isDisplayed())
			{

				WebElement department_name_input_field = driver.findElement(By.xpath("//input[@id='DepartmentName']"));
				department_name_input_field.sendKeys("test Department");
				String result = "Pass";
				System.out.println("This test case is passed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
				driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			}
			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
			}

			System.out.println("10th Test Case");
			row = sheet.createRow(rowCounter);
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to check whether the Add Department button is clickable in the Add edpartment dialog box");

			System.out.println("Test case to check whether the Add Department button is clickable in the Add edpartment dialog box");
			if(driver.findElement(By.xpath("//button[@type='submit' and contains(text(),'Add Department')]")).isDisplayed())
			{
				WebElement add_department_button2 = driver.findElement(By.xpath("//button[@type='submit' and contains(text(),'Add Department')]"));
				add_department_button2.click();
				String result = "Pass";
				System.out.println("This test case is passed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
				driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			}
			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
			}

			System.out.println("11th Test Case");
			row = sheet.createRow(rowCounter);
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to check whether alert pop-ups and the New Department is Added successfully");

			Thread.sleep(1000L);
			System.out.println("Test case to check whether alert pop-ups and the New Department is Added successfully");
			if(driver.findElement(By.xpath("//div[@role='alertdialog']")).getText().equalsIgnoreCase("Added Successfully!\nx"))
			{
				System.out.println("New Department has been added successfully");
				String result = "Pass";
				System.out.println("This test case is passed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
				driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			}
			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
			}

			System.out.println("12th Test Case");
			row = sheet.createRow(rowCounter);
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to check whether the Department Close button is clickable");

			System.out.println("Test case to check whether the Department Close button is clickable");
			if(driver.findElement(By.xpath("//button[@type='button' and contains(text(),'Close')]")).isDisplayed())
			{
				WebElement add_department_close_button = driver.findElement(By.xpath("//button[@type='button' and contains(text(),'Close')]"));
				add_department_close_button.click();
				String result = "Pass";
				System.out.println("This test case is passed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
				driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			}
			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
			}

			System.out.println("13th Test Case");
			row = sheet.createRow(rowCounter);
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to check whether the Add Department dialog box Closes successfully");

			System.out.println("Test case to check whether the Add Department dialog box Closes successfully");
			if(driver.findElement(By.xpath(".//th[contains(text(),'DepartmentName')]")).getText().contains("DepartmentName"))
			{
				String result = "Pass";
				System.out.println("This test case is passed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
				driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			}
			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
			}

			WebElement department_button2 = driver.findElement(By.xpath("//a[contains(text(),'Department')]"));
			JavascriptExecutor jscript = (JavascriptExecutor) driver;
			//driver.manage().timeouts().implicitlyWait(20000, TimeUnit.SECONDS);
			//jscript.executeScript("window.scrollTo(0, -document.body.scrollHeight)");
			driver.manage().timeouts().implicitlyWait(20000, TimeUnit.SECONDS);
			//driver.findElement(By.cssSelector("html")).sendKeys(Keys.CONTROL, Keys.HOME);
			Robot robot = new Robot();
			robot.keyPress(KeyEvent.VK_CONTROL);
			robot.keyPress(KeyEvent.VK_HOME);
			robot.keyRelease(KeyEvent.VK_HOME);
			robot.keyRelease(KeyEvent.VK_CONTROL);
			WebElement home_button = driver.findElement(By.xpath("//a[contains(text(),'Home')]"));
			home_button.click();
			driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			department_button2.click();

			//Deleting a department

			System.out.println("14th Test Case");
			row = sheet.createRow(rowCounter);
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to check whether the delete button is displayed and clickable");

			System.out.println("Test case to check whether the delete button is displayed and clickable");
			if(driver.findElement(By.xpath("//button[@type='button' and contains(text(),'Delete')]")).isDisplayed())
			{
				WebElement add_department_delete_button = driver.findElement(By.xpath("//button[@type='button' and contains(text(),'Delete')]"));
				add_department_delete_button.click();
				String result = "Pass";
				System.out.println("This test case is passed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
				driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			}
			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
			}

			System.out.println("15th Test Case");
			row = sheet.createRow(rowCounter);
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to check whether alert pop-ups");

			System.out.println("Test case to check whether alert pop-ups");
			if(driver.switchTo().alert().getText().equalsIgnoreCase("Are you sure?"))
			{
				driver.switchTo().alert().accept();
				String result = "Pass";
				System.out.println("This test case is passed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
				driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			}
			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
			}

			driver.manage().timeouts().implicitlyWait(5000, TimeUnit.SECONDS);
			//WebElement home_button2 = driver.findElement(By.xpath("//a[contains(text(),'Home')]"));
			//driver.findElement(By.cssSelector("html")).sendKeys(Keys.CONTROL, Keys.END);
			jscript.executeScript("window.scrollTo(0, document.body.scrollHeight)");
			WebElement home_button3 = driver.findElement(By.xpath("//a[contains(text(),'Home')]"));
			home_button3.click();
			driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			department_button2.click();

			//Editing a department

			System.out.println("16th Test Case");
			row = sheet.createRow(rowCounter);
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to check whether the edit button is displayed");

			System.out.println("Test case to check whether the edit button is displayed");
			if(driver.findElement(By.xpath("//button[@type='button' and contains(text(),'Edit')]")).isDisplayed())
			{
				WebElement add_department_edit_button = driver.findElement(By.xpath("//button[@type='button' and contains(text(),'Edit')]"));
				add_department_edit_button.click();
				if(driver.findElement(By.xpath("//div[@class='modal-title h4' and contains(text(),'Edit Department')]")).isDisplayed())
				{
					String result = "Pass";
					System.out.println("This test case is passed");
					row.createCell(2).setCellValue(result);
					rowCounter++;
					driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
				}

			}
			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
			}



			Actions act = new Actions(driver);

			act.sendKeys(Keys.TAB).build().perform();
			act.sendKeys(Keys.TAB).build().perform();

			//Clearing the old department name
			act.sendKeys(Keys.BACK_SPACE).build().perform();
			act.sendKeys("test1").build().perform();

			//act.sendKeys(Keys.RETURN).build().perform();
			// WebElement department_edit_text_input = driver.findElement(By.xpath("//input[@type='text' or @id='DepartmentName']"));
			//WebDriverWait wait = new WebDriverWait(driver, 100);
			// wait.until(ExpectedConditions.visibilityOf(department_edit_text_input));
			//  department_edit_text_input.click();
			// department_edit_text_input.clear();
			// department_edit_text_input.sendKeys("test1");
			driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);

			//Updating a department
			//WebElement department_edit_update_button = driver.findElement(By.xpath("//button[@type='submit' and contains(text(),'Update Department')]"));
			//department_edit_update_button.click();

//      String testCaseDesc = "Test case to check whether the Update department button is displayed and clickable";
//      System.out.println(testCaseDesc);

			System.out.println("17th Test Case");
			row = sheet.createRow(rowCounter);
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to check whether the Update department button is displayed and clickable");

			System.out.println("Test case to check whether the Update department button is displayed and clickable");
			if(driver.findElement(By.xpath("//button[@type='submit' and contains(text(),'Update Department')]")).isDisplayed())
			{
				WebElement element = driver.findElement(By.xpath("//button[@type='submit' and contains(text(),'Update Department')]"));
				act.sendKeys(Keys.TAB).build().perform();
				// driver.findElement(By.xpath("//button[@type='submit' and contains(text(),'Update Department')]")).click();
				//act.click().build().perform();
				JavascriptExecutor jse = (JavascriptExecutor)driver;
				jse.executeScript("arguments[0].click()", element);
				String result = "Pass";
				System.out.println("This test case is passed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
				driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			}
			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;

			}

			System.out.println("18th Test Case");
			row = sheet.createRow(rowCounter);
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to check whether alert pop-ups and the Department is Updated successfully");

			System.out.println("Test case to check whether alert pop-ups and the Department is Updated successfully");
			if(driver.findElement(By.xpath("//div[@role='alertdialog']")).getText().equalsIgnoreCase("Updated Successfully!\nx"))
			{
				System.out.println("Department has been updated successfully");
				String result = "Pass";
				System.out.println("This test case is passed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
				driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			}
			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
			}

			System.out.println("19th Test Case");
			row = sheet.createRow(rowCounter);
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to check whether the CLose button is dsiplayed and clickable");

			System.out.println("Test case to check whether the CLose button is dsiplayed and clickable");
			if(driver.findElement(By.xpath("//button[@type='button' and contains(text(),'Close')]")).isDisplayed())
			{
				WebElement element = driver.findElement(By.xpath("//button[@type='button' and contains(text(),'Close')]"));
				act.sendKeys(Keys.TAB).build().perform();
				JavascriptExecutor jse = (JavascriptExecutor)driver;
				jse.executeScript("arguments[0].click()", element);
				String result = "Pass";
				System.out.println("This test case is passed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
				driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);;
			}
			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
			}

			driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);

			//Clicking on employee tab

			System.out.println("20th Test Case");
			row = sheet.createRow(rowCounter);
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to check whether the Employee tab is displayed");

			System.out.println("Test case to check whether the Employee tab is displayed");
			if(driver.findElement(By.xpath("//a[contains(text(),'Employee')]")).isDisplayed())
			{
				WebElement employee_button2 = driver.findElement(By.xpath("//a[contains(text(),'Employee')]"));
				employee_button2.click();
				driver.manage().timeouts().implicitlyWait(5000, TimeUnit.SECONDS);

				js.executeScript("window.scrollTo(0, document.body.scrollHeight)");
				String result = "Pass";
				System.out.println("This test case is passed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
				driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			}
			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
			}


			// Clicking on Add employee button

			System.out.println("21st Test Case");
			row = sheet.createRow(rowCounter);
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to check whether Add employee button is displayed and clickable");

			System.out.println("Test case to check whether Add employee button is displayed and clickable");
			if(driver.findElement(By.xpath("//button[contains(text(),'Add Employee')]")).isDisplayed())
			{
				WebElement add_employee_button = driver.findElement(By.xpath("//button[contains(text(),'Add Employee')]"));
				add_employee_button.click();
				String result = "Pass";
				System.out.println("This test case is passed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
				//driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
				Thread.sleep(100L);
			}
			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
			}

			driver.manage().timeouts().implicitlyWait(5000, TimeUnit.SECONDS);
			// Adding New Employee
			WebElement employee_name_input_field = driver.findElement(By.xpath("//input[@id='EmployeeName']"));
			employee_name_input_field.sendKeys("employee1");
			driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);

			Select department = new Select(driver.findElement(By.id("Department")));
			//department.selectByVisibleText("test Department");
			department.selectByIndex(1);

			WebElement employee_email_id_input_field = driver.findElement(By.xpath("//input[@id='MailID']"));
			employee_email_id_input_field.sendKeys("xyz@gmail.com");
			driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);

			// Selecting date

			//Adding an employee

			Actions act2 = new Actions(driver);

			//for (int i=0; i < 4; i++)
			//{
			//	 act2.sendKeys(Keys.TAB).build().perform();
			// }

			System.out.println("22nd Test Case");
			row = sheet.createRow(rowCounter);
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to check whether the date picker is displayed and clickable");

			System.out.println("Test case to check whether the date picker is displayed and clickable");
			if(driver.findElement(By.xpath("//input[@type='date' and @id='DOJ']")).isDisplayed())
			{
				JavascriptExecutor jse = (JavascriptExecutor)driver;
				jse.executeScript("dt = new Date()");
				jse.executeScript("date = dt.getFullYear() + '-' + (((dt.getMonth() + 1) < 10) ? '0' : '') + (dt.getMonth() + 1) + '-' + ((dt.getDate() < 10) ? '0' : '') + dt.getDate()");
				jse.executeScript("document.getElementById(\"DOJ\").value = date");
				String result = "Pass";
				System.out.println("This test case is passed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
				driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			}
			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
			}


//	 dt = new Date();
//	 date = dt.getFullYear() + '-' + (((dt.getMonth() + 1) < 10) ? '0' : '') + (dt.getMonth() + 1) + '-' + ((dt.getDate() < 10) ? '0' : '') + dt.getDate();
//	 document.getElementById("DOJ").value = date;



			System.out.println("23rd Test Case");
			row = sheet.createRow(rowCounter);
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to check whether the add employee button is displayed in the Add employee dialog box");

			System.out.println("Test case to check whether the add employee button is displayed in the Add employee dialog box");
			if(driver.findElement(By.xpath("//button[@type='submit' and contains(text(),'Add Employee')]")).isDisplayed())
			{
				WebElement button = driver.findElement(By.xpath("//button[@type='submit' and contains(text(),'Add Employee')]"));
				button.click();
				String result = "Pass";
				System.out.println("This test case is passed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
				driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			}
			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
			}

			System.out.println("24th Test Case");
			row = sheet.createRow(rowCounter);
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to check whether alert pop-ups and the New Employee is Added successfully");

			System.out.println("Test case to check whether alert pop-ups and the New Employee is Added successfully");
			if(driver.findElement(By.xpath("//div[@role='alertdialog']")).getText().equalsIgnoreCase("Added Successfully\nx"))
			{
				System.out.println("New Employee has been added successfully");
				System.out.println("This test case is passed");
			}
			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
			}

			// act2.sendKeys(Keys.TAB).build().perform();

			System.out.println("25th Test Case");
			row = sheet.createRow(rowCounter);
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to check whether the Close button is displayed in the Add employee dialog box");

			System.out.println("Test case to check whether the Close button is displayed in the Add employee dialog box");

			System.out.println("-----1-----");
			// if(driver.findElement(By.xpath("//button[@type='button' and contains(text(),'Close')]")).isDisplayed())
			// {
			if(driver.findElement(By.xpath("//button[@type='button' and @class='close']")).isDisplayed())
			{
				System.out.println("-----2-----");
				//WebElement button = driver.findElement(By.xpath("//button[@type='button' and contains(text(),'Close')]"));
				WebElement button = driver.findElement(By.xpath("//button[@type='button' and @class='close']"));
				button.click();

				act2.moveToElement(driver.findElement(By.xpath("//div[@id='root']"))).click().perform();

				// JavascriptExecutor jse = (JavascriptExecutor)driver;
//    	 jse.executeScript("document.getElementsByName('C')[0].focus();");
				System.out.println("This test case is passed");
				driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
//    	 Thread.sleep(2000L);
				WebElement department_button3 = driver.findElement(By.xpath("//a[contains(text(),'Department')]"));
				department_button3.click();
				driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
				WebElement employee_button2 = driver.findElement(By.xpath("//a[contains(text(),'Employee')]"));
				employee_button2.click();
			}
			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
			}

			System.out.println("-----3-----");

			// Editing an employee

			System.out.println("26th Test Case");
			row = sheet.createRow(rowCounter);
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to check whether the Edit button is displayed under the Employee tab");

			System.out.println("Test case to check whether the Edit button is displayed under the Employee tab");

			if(driver.findElement(By.xpath("//button[@type='button' and contains(text(),'Edit')]")).isDisplayed())
			{

				WebElement button = driver.findElement(By.xpath("//button[@type='button' and contains(text(),'Edit')]"));
				button.click();
				//  WebElement add_department_edit_button = driver.findElement(By.tagName("script"));
				//  String htmlCode = (String) ((JavascriptExecutor) driver).executeScript("return arguments[0].innerHTML;", add_department_edit_button);
				//add_department_edit_button.click();
				//Actions act3 = new Actions(driver);

				//  for (int i=0; i < 7; i++)
				//  {
				// 	 act3.sendKeys(Keys.TAB).build().perform();
				//  }

				//  act3.click();
				driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
				if(driver.findElement(By.xpath("//div[@class='modal-title h4' and contains(text(),'Edit Employee')]")).isDisplayed())
				{
					driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
					String result = "Pass";
					System.out.println("This test case is passed");
					row.createCell(2).setCellValue(result);
					rowCounter++;
					driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
				}

			}
			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
			}

			// Updating New Employee

			System.out.println("27th Test Case");
			row = sheet.createRow(rowCounter);
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to to whether the employee name input field is displayed");

			System.out.println("Test case to to whether the employee name input field is displayed");
			if(driver.findElement(By.xpath("//input[@id='EmployeeName']")).isDisplayed())
			{
				WebElement edit_employee_name_input_field = driver.findElement(By.xpath("//input[@id='EmployeeName']"));
//    	 edit_employee_name_input_field.click();
				// act.sendKeys(Keys.TAB).build().perform();
				edit_employee_name_input_field.clear();
//
//         edit_employee_name_input_field.sendKeys("u'\ue009' + u'\ue003'");
				// edit_employee_name_input_field.sendKeys(Keys.CONTROL + "a");
				// edit_employee_name_input_field.sendKeys(Keys.DELETE);

				// driver.findElement(By.id("EmployeeName")).clear();
				JavascriptExecutor jse = (JavascriptExecutor)driver;
				jse.executeScript("document.getElementById(\"EmployeeName\").value = \"John\"");
				//document.getElementById("EmployeeName").value = "John";

//         edit_employee_name_input_field.sendKeys("John");
				String result = "Pass";
				System.out.println("This test case is passed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
				driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			}
			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
			}

			System.out.println("28th Test Case");
			row = sheet.createRow(rowCounter);
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to to whether the department input field is displayed");

			System.out.println("Test case to to whether the department input field is displayed");
			if(driver.findElement(By.id("Department")).isDisplayed())
			{
				Select edit_department = new Select(driver.findElement(By.id("Department")));
				//department.selectByVisibleText("test Department");
				edit_department.selectByIndex(1);
				String result = "Pass";
				System.out.println("This test case is passed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
				driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			}
			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
			}

			System.out.println("29th Test Case");
			row = sheet.createRow(rowCounter);
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to to whether the mailID input field is displayed");

			System.out.println("Test case to to whether the mailID input field is displayed");
			if(driver.findElement(By.xpath("//input[@id='MailID']")).isDisplayed())
			{
				WebElement edit_employee_email_id_input_field = driver.findElement(By.xpath("//input[@id='MailID']"));
				edit_employee_email_id_input_field.clear();
				edit_employee_email_id_input_field.sendKeys("xyz@gmail.com");
				String result = "Pass";
				System.out.println("This test case is passed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
				driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			}
			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
			}

			// Selecting date

			//Editing an employee

			//  Actions act4 = new Actions(driver);

			// for (int i=0; i < 3; i++)
			// {
			// act4.sendKeys(Keys.TAB).build().perform();

			// }

			System.out.println("30th Test Case");
			row = sheet.createRow(rowCounter);
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to to whether the update employee button is displayed and clickable");

			System.out.println("Test case to to whether the update employee button is displayed and clickable");
			if(driver.findElement(By.xpath("//body/div[9]/div[1]/div[1]/div[2]/div[1]/div[1]/form[1]/div[6]/button[1]")).isDisplayed())
			{
				//WebElement button = driver.findElement(By.xpath("//button[@type='submit' and contains(text(),'Update Employee')]"));
				//WebDriverWait wait = WebDriverWait(driver,100);
//    	WebElement  element=driver.findElement(By.xpath("//button[@type='submit' and contains(text(),'Update Employee')]"));
				WebElement  element=driver.findElement(By.xpath("//body/div[9]/div[1]/div[1]/div[2]/div[1]/div[1]/form[1]/div[6]/button[1]"));
				//Create object of Robot class
				// Robot rb = new Robot();

				//Find x and y coordinates to pass to mouseMove method
				//1. Get the size of the current window.
				//2. Dimension class is similar to java Point class which represents a location in a two-dimensional (x, y) coordinate space.
				//But here Point point = element.getLocation() method can't be used to find the position
				//as this is Windows Popup and its locator is not identifiable using browser developer tool
				// Dimension i = driver.manage().window().getSize();
				// System.out.println("Dimension x and y :"+i.getWidth()+" "+i.getHeight());
				//3. Get the height and width of the screen
				// int x = (i.getWidth()/4)+20;
				// int y = (i.getHeight()/10)+50;
				//4. Now, adjust the x and y coordinates with reference to the Windows popup size on the screen
				//e.g. On current screen , Windows popup displays on almost 1/4th of the screen . So with reference to the same, file name x and y position is specified.
				//Note : Please note that coordinates calculated in this sample i.e. x and y may vary as per the screen resolution settings
				// rb.mouseMove(x,y);

				//Clicks Left mouse button
				// rb.mousePress(InputEvent.BUTTON1_DOWN_MASK);
				// rb.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
				// System.out.println("Browse button clicked");
				// Thread.sleep(2000);

				//Closes the Desktop Windows popup
				// rb.keyPress(KeyEvent.VK_ENTER);
				// System.out.println("Closed the windows popup");
				// Thread.sleep(1000);

				JavascriptExecutor ex=(JavascriptExecutor)driver;
				ex.executeScript("arguments[0].click()", element);
				String result = "Pass";
				System.out.println("This test case is passed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
				driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			}
			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
			}

			System.out.println("31st Test Case");
			row = sheet.createRow(rowCounter);
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to check whether alert pop-ups and the Employee is Updated successfully");

			System.out.println("Test case to check whether alert pop-ups and the Employee is Updated successfully");
			if(driver.findElement(By.xpath("//div[@role='alertdialog']")).getText().equalsIgnoreCase("Updated Successfully\nx"))
			{
				System.out.println("The Employee has been updated successfully");
				String result = "Pass";
				System.out.println("This test case is passed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
				driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			}
			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
			}
			act2.sendKeys(Keys.TAB).build().perform();

			System.out.println("32nd Test Case");
			row = sheet.createRow(rowCounter);
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to to whether the Close button is displayed and clickable");

			System.out.println("Test case to to whether the Close button is displayed and clickable");
			if(driver.findElement(By.xpath("//button[@type='button' and contains(text(),'Close')]")).isDisplayed())
			{
				WebElement element = driver.findElement(By.xpath("//button[@type='button' and contains(text(),'Close')]"));
				JavascriptExecutor ex=(JavascriptExecutor)driver;
				ex.executeScript("arguments[0].click()", element);
				String result = "Pass";
				System.out.println("This test case is passed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
				driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			}
			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
			}

			// Deleting an Employee

			System.out.println("33rd Test Case");
			row = sheet.createRow(rowCounter);
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to check whether the delete button is displayed and clickable");

			System.out.println("Test case to check whether the delete button is displayed and clickable");
			if(driver.findElement(By.xpath("//button[@type='button' and contains(text(),'Delete')]")).isDisplayed())
			{
				WebElement add_department_delete_button = driver.findElement(By.xpath("//button[@type='button' and contains(text(),'Delete')]"));
				add_department_delete_button.click();
				String result = "Pass";
				System.out.println("This test case is passed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
				driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			}
			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
			}

			System.out.println("34th Test Case");
			row = sheet.createRow(rowCounter);
			row.createCell(0).setCellValue(rowCounter);
			row.createCell(1).setCellValue("Test case to check whether alert pop-ups");

			System.out.println("Test case to check whether alert pop-ups");
			if(driver.switchTo().alert().getText().equalsIgnoreCase("Are you sure?"))
			{
				driver.switchTo().alert().accept();
				String result = "Pass";
				System.out.println("This test case is passed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
				driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			}
			else
			{
				String result = "Fail";
				System.out.println("This test case is failed");
				row.createCell(2).setCellValue(result);
				rowCounter++;
			}

			driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			WebElement department_button3 = driver.findElement(By.xpath("//a[contains(text(),'Department')]"));
			department_button3.click();
			driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			WebElement employee_button2 = driver.findElement(By.xpath("//a[contains(text(),'Employee')]"));
			employee_button2.click();

		}
		catch(Exception e) {
			System.out.println("Exception occurred!");
			e.printStackTrace();
		}
		finally {

			FileOutputStream fileOut = new FileOutputStream(filename);
			workbook.write(fileOut);
			//closing the Stream
			fileOut.close();
			//closing the workbook
			workbook.close();
			System.out.println("File created successfully!");
		}

	}


	//}
	@AfterMethod
	public void tearDown() {

		//driver.close();
		// driver.quit();
	}

}
