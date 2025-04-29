package com.twowheeler.base;

import com.twowheeler.utilities.ExcelReader;
import com.twowheeler.utilities.MonitoringMail;
import com.twowheeler.utilities.TestConfig;

import io.github.bonigarcia.wdm.WebDriverManager;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

import javax.mail.MessagingException;
import javax.mail.internet.AddressException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeSuite;
 
public class TestBaseNew {
	public static WebDriver driver;
	public static Properties baseconf = new Properties();
	public static Properties baseloc = new Properties();
	public static Properties loc = new Properties();
	public static  FileReader basedata;
	public static  FileReader loginlogout;
	public static  FileReader Locators;
	static String messageBody;
	static String testDataPath;
	
	public static ExcelReader excel = new ExcelReader(
			System.getProperty("user.dir") + "\\src\\test\\resources\\TestData\\PrivateCar.xlsx");

	
	@BeforeSuite(alwaysRun = true)
	public void setup() throws IOException, InterruptedException {
		
		if(driver==null) {
			  basedata = new FileReader(System.getProperty("user.dir") + "\\src\\test\\resources\\ConfigFiles\\ZNew.Base.Config.Properties"); 
			  loginlogout = new FileReader(System.getProperty("user.dir")+ "\\src\\test\\resources\\ConfigFiles\\Z.Base.Locator.Properties");
			  Locators = new FileReader(System.getProperty("user.dir")+ "\\src\\test\\resources\\ConfigFiles\\New.Locator.Properties");
					  
			  baseconf.load(basedata);
			  baseloc.load(loginlogout);
			  loc.load(Locators);
		  }
		  if(baseconf.getProperty("browser").equalsIgnoreCase("ch")) {
			  WebDriverManager.chromedriver().setup();
			// Create a path for Downloads folder inside your project
		        String downloadFilepath = System.getProperty("user.dir") + File.separator +  "QuoteLetter";

		        // Set Chrome preferences
		        Map<String, Object> prefs = new HashMap<>();
		        prefs.put("download.default_directory", downloadFilepath); // Set your custom download path
		        prefs.put("plugins.always_open_pdf_externally", true);     // Disable PDF viewer in Chrome

		        ChromeOptions options = new ChromeOptions();
		        options.setExperimentalOption("prefs", prefs);

		        // Launch Chrome with options
				driver = new ChromeDriver(options);
				driver.manage().window().maximize();
				driver.get(baseconf.getProperty("testurl"));
				System.out.println("URL opened successfully");
				Thread.sleep(5000);
				driver.findElement(By.id(baseloc.getProperty("unm"))).sendKeys(baseconf.getProperty("agentID"));
				Thread.sleep(1000);
				System.out.println("Username Entered successfully");
				Thread.sleep(1000);
				driver.findElement(By.id(baseloc.getProperty("psd"))).sendKeys(baseconf.getProperty("password"));
				System.out.println("Password Entered successfully");
				Thread.sleep(2000);
				driver.findElement(By.xpath(baseloc.getProperty("Login"))).click();
				System.out.println("Login Button clicked successfully");
				System.out.println("User log in successfully");
				Thread.sleep(6000);
				System.out.println("Home Page opened successfully");

		  }
		  else if(baseconf.getProperty("browser").equalsIgnoreCase("ff")) {
			  WebDriverManager.firefoxdriver().setup();
			  driver = new FirefoxDriver(); 
			  driver.manage().window().maximize();
			  driver.get(baseconf.getProperty("testurl"));
				System.out.println("URL opened successfully");
				Thread.sleep(5000);
				driver.findElement(By.id(baseloc.getProperty("unm"))).sendKeys(baseconf.getProperty("agentID"));
				Thread.sleep(1000);
				System.out.println("Username Entered successfully");
				Thread.sleep(1000);
				driver.findElement(By.id(baseloc.getProperty("Psd"))).sendKeys(baseconf.getProperty("password"));
				System.out.println("Password Entered successfully");
				Thread.sleep(2000);
				driver.findElement(By.xpath(baseloc.getProperty("Login"))).click();
				System.out.println("Login Button clicked successfully");
				System.out.println("User log in successfully");
				Thread.sleep(6000);
				System.out.println("Home Page opened successfully");
		  }
		  else if(baseconf.getProperty("browser").equalsIgnoreCase("ms")) {
			  WebDriverManager.edgedriver().setup();
			  driver = new EdgeDriver();
			  driver.manage().window().maximize();
			  driver.get(baseconf.getProperty("testurl"));
				System.out.println("URL opened successfully");
				Thread.sleep(5000);	
				driver.findElement(By.id(baseloc.getProperty("unm"))).sendKeys(baseconf.getProperty("agentID"));
				Thread.sleep(1000);
				System.out.println("Username Entered successfully");
				Thread.sleep(1000);
				driver.findElement(By.id(baseloc.getProperty("psd"))).sendKeys(baseconf.getProperty("password"));
				System.out.println("Password Entered successfully");
				Thread.sleep(2000);
				driver.findElement(By.xpath(baseloc.getProperty("Login"))).click();
				System.out.println("Login Button clicked successfully");
				System.out.println("User log in successfully");
				Thread.sleep(6000);
				System.out.println("Home Page opened successfully");
		  }
	  }
	public void Baseclick(String locator) throws InterruptedException {

		if (locator.endsWith("_CSS")) {
			driver.findElement(By.cssSelector(baseloc.getProperty(locator))).click();
		} else if (locator.endsWith("_XPATH")) {
			Thread.sleep(10000);
			driver.findElement(By.xpath(baseloc.getProperty(locator))).click();
		} else if (locator.endsWith("_ID")) {
			driver.findElement(By.id(baseloc.getProperty(locator))).click();
		}
		//CustomListeners.testReport.get().log(Status.INFO, "Clicking on : " + locator);
	}
	public void click(String locator) throws InterruptedException {

		if (locator.endsWith("_CSS")) {
			driver.findElement(By.cssSelector(loc.getProperty(locator))).click();
		} else if (locator.endsWith("_XPATH")) {
			Thread.sleep(10000);
			driver.findElement(By.xpath(loc.getProperty(locator))).click();
		} else if (locator.endsWith("_ID")) {
			driver.findElement(By.id(loc.getProperty(locator))).click();
		}
		//CustomListeners.testReport.get().log(Status.INFO, "Clicking on : " + locator);
	}
	public void Enter(String locator, String value) throws InterruptedException {

		if (locator.endsWith("_CSS")) {
			driver.findElement(By.cssSelector(loc.getProperty(locator))).sendKeys(value);
		} else if (locator.endsWith("_XPATH")) {
			driver.findElement(By.xpath(loc.getProperty(locator))).sendKeys(value);
		} else if (locator.endsWith("_ID")) {
			driver.findElement(By.id(loc.getProperty(locator))).sendKeys(value);
		}

		//CustomListeners.testReport.get().log(Status.INFO, "Typing in : " + locator + " entered value as " + value);

	}
	public void Clear(String locator) throws InterruptedException {

		if (locator.endsWith("_CSS")) {
			driver.findElement(By.cssSelector(loc.getProperty(locator))).clear();
		} else if (locator.endsWith("_XPATH")) {
			Thread.sleep(10000);
			driver.findElement(By.xpath(loc.getProperty(locator))).clear();
		} else if (locator.endsWith("_ID")) {
			driver.findElement(By.id(loc.getProperty(locator))).clear();
		}
		//CustomListeners.testReport.get().log(Status.INFO, "Clicking on : " + locator);
	}

	
	static WebElement dropdown;

	public void select(String locator, String value) {

		if (locator.endsWith("_CSS")) {
			dropdown = driver.findElement(By.cssSelector(loc.getProperty(locator)));
		} else if (locator.endsWith("_XPATH")) {
			dropdown = driver.findElement(By.xpath(loc.getProperty(locator)));
		} else if (locator.endsWith("_ID")) {
			dropdown = driver.findElement(By.id(loc.getProperty(locator)));
		}
		
		Select select = new Select(dropdown);
		select.selectByVisibleText(value);

		//CustomListeners.testReport.get().log(Status.INFO, "Selecting from dropdown : " + locator + " value as " + value);

	}
	public void output(String value){
		System.out.println(value);	
	}

	public String print(String locator) throws InterruptedException {
		String element = null;
		if (locator.endsWith("_CSS")) {
			element = driver.findElement(By.cssSelector(loc.getProperty(locator))).getAttribute("value");
		} else if (locator.endsWith("_XPATH")) {
			element = driver.findElement(By.xpath(loc.getProperty(locator))).getAttribute("value");
		} else if (locator.endsWith("_ID")) {
			element = driver.findElement(By.id(loc.getProperty(locator))).getAttribute("value");
		}
		// CustomListeners.testReport.get().log(Status.INFO, "Clicking on : " +
		// locator);
		return element;

	}
	public void clear(String locator) throws InterruptedException {
		 
		if (locator.endsWith("_CSS")) {
			driver.findElement(By.cssSelector(loc.getProperty(locator))).clear();
		} else if (locator.endsWith("_XPATH")) {
			Thread.sleep(5000);
			driver.findElement(By.xpath(loc.getProperty(locator))).clear();
		} else if (locator.endsWith("_ID")) {
			driver.findElement(By.id(loc.getProperty(locator))).clear();
		}
		//CustomListeners.testReport.get().log(Status.INFO, "Clicking on : " + locator);
	}
	public void write(String result,Method m, int colNum, int rowNum) {
		String excelFilePath = System.getProperty("user.dir") + "\\src\\test\\resources\\TestData\\PrivateCar.xlsx";
		try {
		FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
		Workbook workbook = WorkbookFactory.create(inputStream);
		
		Sheet sheet = workbook.getSheet(m.getName());
		Cell cell2Update = sheet.getRow(rowNum).getCell(colNum);
		cell2Update.setCellValue(result);
		inputStream.close();
		
		FileOutputStream outputStream = new FileOutputStream(
		System.getProperty("user.dir") + "\\src\\test\\resources\\TestData\\PrivateCar.xlsx");
		workbook.write(outputStream);
		workbook.close();
		outputStream.close();
		
		} catch (Exception ex) {
		ex.printStackTrace();
		}
		}
	public String read(String sheetname, int colNum, int rowNum) {
		String excelFilePath = System.getProperty("user.dir")
				+ "\\src\\test\\resources\\TestData\\PrivateCar.xlsx";
		try {
			FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
			Workbook workbook = WorkbookFactory.create(inputStream);
			String data = "";
			if (workbook.getSheet(sheetname).getRow(rowNum).getCell(colNum).getCellType() == CellType.NUMERIC) {
				int celldata = (int) workbook.getSheet(sheetname).getRow(rowNum).getCell(colNum).getNumericCellValue();
				data = String.valueOf(celldata);
			} else {
				data = workbook.getSheet(sheetname).getRow(rowNum).getCell(colNum).getStringCellValue();
			}
			return data;
		} catch (Exception ex) {
			ex.printStackTrace();
		}
		return excelFilePath;
	}
	
	@AfterSuite(alwaysRun = true)
	  public void tearDown() throws InterruptedException {
		MonitoringMail mail = new MonitoringMail();

		messageBody = "Dear All,<br><br>" +                        // "Dear All" followed by two line breaks
		        "Please find the attached Test Results.<br><br>"; // Hyperlinked URL
		// Path to the test data file
		testDataPath = System.getProperty("user.dir") + "\\src\\test\\resources\\TestData\\PrivateCar.xlsx";

		try {
			//String timestamp = new SimpleDateFormat("yyyy.MM.dd.HH.mm.ss").format(new Date());
			//String repName = "Test-Report-"+timestamp+".html";
			String reportPath = System.getProperty("user.dir") + "\\reports\\PCNV-UAT-Reports.html";
			mail.sendMail(TestConfig.server, TestConfig.from, TestConfig.to, TestConfig.subject, messageBody, testDataPath, reportPath);
		} catch (AddressException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (MessagingException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
			Thread.sleep(5000);
			driver.findElement(By.xpath(baseloc.getProperty("profile"))).click();
			System.out.println("Profile Clicked successfully");
			Thread.sleep(5000);
			driver.findElement(By.xpath(baseloc.getProperty("Sobtn"))).click();
			System.out.println("Sign Out clicked successfully");
			System.out.println("User log Out successfully");
			Thread.sleep(3000);
			driver.quit();
		    System.out.println("Browser Closed successfully");
	  }
		
	}
