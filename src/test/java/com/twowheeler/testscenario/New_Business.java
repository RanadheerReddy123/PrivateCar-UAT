package com.twowheeler.testscenario;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.Method;
import java.net.InetAddress;
import java.net.UnknownHostException;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.Hashtable;

import javax.mail.MessagingException;
import javax.mail.internet.AddressException;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.Select;
import org.testng.ITestResult;
import org.testng.SkipException;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.MediaEntityBuilder;
import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.markuputils.ExtentColor;
import com.aventstack.extentreports.markuputils.Markup;
import com.aventstack.extentreports.markuputils.MarkupHelper;
import com.aventstack.extentreports.reporter.ExtentSparkReporter;
import com.aventstack.extentreports.reporter.configuration.Theme;
import com.twowheeler.base.TestBase;
import com.twowheeler.base.TestBaseNew;
import com.twowheeler.utilities.MonitoringMail;
import com.twowheeler.utilities.TestConfig;
import com.twowheeler.utilities.TestUtil;

import Utilities_UAT_ScreenRecord.Motor_PersonBP_ScreenRecord;

public class New_Business extends TestBaseNew {

	public ExtentSparkReporter htmlReporter;
	public ExtentReports extent;
	public ExtentTest test;
	public JavascriptExecutor js;
	private int rowNum = 1;
	static String messageBody;
	static String testDataPath;
	static double finalSI;
	String repName;

	@BeforeTest(alwaysRun = true)
	public void setReport() throws IOException {
		htmlReporter = new ExtentSparkReporter("./reports/PCNV-UAT-Reports.html");

		htmlReporter.config().setEncoding("utf-8");
		htmlReporter.config().setDocumentTitle("Private Car NV Automation Test Reports");
		htmlReporter.config().setReportName("Private Car NV Automation Test Results");
		htmlReporter.loadXMLConfig(new File(".\\ExtentReportXMLs\\CustomeHeader.xml"));
		htmlReporter.config().setTheme(Theme.STANDARD);

		extent = new ExtentReports();
		extent.attachReporter(htmlReporter);

		extent.setSystemInfo("Automation Tester", "Ranadheer Reddy");
		extent.setSystemInfo("Organisation", "Serole Technologies");
	}

	@Test(dataProviderClass = TestUtil.class, dataProvider = "dp", alwaysRun = true)
	public void NewVehicle(Hashtable<String, String> data, Method m) throws Exception {
		// ScreenRecord Start
		Motor_PersonBP_ScreenRecord.startRecord("PersonBP");
		try {
			test = extent.createTest("Private Car Test");
			if (data.get("Runmode").equalsIgnoreCase("N")) {
				throw new SkipException("Skipping the test case " + "pcnv".toUpperCase() + " as the Run mode is NO");
			}else {
				test.pass("<b>" + "<font color=" + "green>" + "Runmode is satisfied" + "</font>" + "</b>");
			}
			// Clicking on QMS Tile
			//Thread.sleep(5000);
			Home("QMSTile_XPATH");
			output("Qms Tile Click successfully");
			//Thread.sleep(5000);

			// Quote Dash board display
			Baseclick("NewQuote_XPATH");
			output("New Quote Button clicked successfully");
			Baseclick("MotorTile_XPATH");
			output("Motor Tile clicked successfully");
			//Baseclick("RegVechl_XPATH");
			//output("Registered Vehicle Selected Successfully");
			Baseclick("NewMotor_XPATH");
			output("New Vehicle Selected Succesfully");
			Baseclick("nxtbtn_XPATH");
			output("Next Button Clicked successfully");
			if (data.get("Vehicle Type").equalsIgnoreCase("Private Motor Car")) {
				click("PvtCar_XPATH");
			} else if (data.get("Vehicle Type").equalsIgnoreCase("Motorcycle")) {
				click("MotorCycle_XPATH");
			} else if (data.get("Vehicle Type").equalsIgnoreCase("Commercial")) {
				click("Commercial_XPATH");
			}
			output("Vehicle Type is selected successfully");
			click("POUDrpDon2_XPATH");
			if (data.get("Place Of Use").equalsIgnoreCase("Johor")) {
				click("Johor_XPATH");
			} else if (data.get("Place Of Use").equalsIgnoreCase("Kedah")) {
				click("Kedah_XPATH");
			} else if (data.get("Place Of Use").equalsIgnoreCase("Kelantan")) {
				click("Kelantan_XPATH");
			} else if (data.get("Place Of Use").equalsIgnoreCase("Kuala Lumpur")) {
				click("KualaLumpur_XPATH");
			} else if (data.get("Place Of Use").equalsIgnoreCase("Labuan")) {
				click("Labuan_XPATH");
			} else if (data.get("Place Of Use").equalsIgnoreCase("Melaka")) {
				click("Melaka_XPATH");
			} else if (data.get("Place Of Use").equalsIgnoreCase("Negeri Sembilan")) {
				click("NegeriSembilan_XPATH");
			} else if (data.get("Place Of Use").equalsIgnoreCase("Pahang")) {
				click("Pahang_XPATH");
			} else if (data.get("Place Of Use").equalsIgnoreCase("Perak")) {
				click("Perak_XPATH");
			} else if (data.get("Place Of Use").equalsIgnoreCase("Perlis")) {
				click("Perlis_XPATH");
			} else if (data.get("Place Of Use").equalsIgnoreCase("Pulau Pinang")) {
				click("PulauPinang_XPATH");
			} else if (data.get("Place Of Use").equalsIgnoreCase("Putrajaya")) {
				click("Putrajaya_XPATH");
			} else if (data.get("Place Of Use").equalsIgnoreCase("Sabah")) {
				click("Sabah_XPATH");
			} else if (data.get("Place Of Use").equalsIgnoreCase("Sarawak")) {
				click("Sarawak_XPATH");
			} else if (data.get("Place Of Use").equalsIgnoreCase("Selangor")) {
				click("Selangor_XPATH");
			} else if (data.get("Place Of Use").equalsIgnoreCase("Terengganu")) {
				click("Terengganu_XPATH");
			}
			output("Place Of Use selected successfully");
			test.log(Status.INFO, "Place of Use is selected successfully");
			//Thread.sleep(5000);
			WebElement button = driver.findElement(By.xpath("//button[@id='cancel']"));
			js = (JavascriptExecutor) driver;
			js.executeScript("arguments[0].scrollIntoView(true);", button);
			Enter("Engine_XPATH",data.get("Engine"));
			output("Engine value is given successfully");
			Enter("Chassis_XPATH",data.get("Chassis"));
			output("Chassis value is given successfully");
			click("VU_XPATH");
			Enter("VUS_XPATH",data.get("Vehicle Use"));
			click("VUSelect_XPATH");
			output("Vehicle Use selected successfully");
			//Thread.sleep(5000);
			Enter("EngineCpcty_XPATH",data.get("Engine Capacity").replace(".0", ""));
			output("Engine Capacity value is given successfully");
			WebElement button2 = driver.findElement(By.xpath("/html/body/dx-dijta-home-root/dx-vertical-layout/dx-layout/content/app-parent-qms/dx-compact-content/dx-layout-content/div/div/dx-qms-private-motor/div/dx-qms-nav-menu-layout/div[2]/div/div/app-private-motor-vehicle-details/form/div/dx-qms-nav-menu-content/div/div/div/div[1]/div[1]/dx-card/div/div/div/app-vehicle-search/div/div/div/div[4]/dx-qms-edit-vehicle-details/div/div/div/dx-card/div/div[3]"));
			js = (JavascriptExecutor) driver;
			js.executeScript("arguments[0].scrollIntoView(true);", button2);
			click("Make_XPATH");
			Enter("MakeS_XPATH",data.get("Make"));
			click("MakeSelect_XPATH");
			output("Make value is given successfully");
			//Thread.sleep(5000);	
			click("NewModClick_XPATH");
			output("Model Drop Down Clicked successfully");
			Enter("NewModEntr_XPATH", data.get("Model"));
			output("Model Given successfully");
			click("NewModSel_XPATH");
			output("Model Selected successfully");
			click("YOM_XPATH");
			Enter("YOMS_XPATH",data.get("Year").replace(".0", ""));
			click("YOMSelect_XPATH");
			output("Year of Manufacture is given successfully");
			click("Variant_XPATH");
			Enter("VariantS_XPATH",data.get("Variant"));
			click("VariantSelect_XPATH");
			output("Variant value is given successfully");
			Enter("VariantNameS_XPATH",data.get("Variant2"));
			output("Variant Name value is given successfully");
			Enter("Seat_XPATH",data.get("SeatCpcty").replace(".0", ""));
			output("Seating capacity value is given successfully");
			WebElement button3 = driver.findElement(By.xpath("/html/body/dx-dijta-home-root/dx-vertical-layout/dx-layout/content/app-parent-qms/dx-compact-content/dx-layout-content/div/div/dx-qms-private-motor/div/dx-qms-nav-menu-layout/div[2]/div/div/app-private-motor-vehicle-details/form/div/dx-qms-nav-menu-content/div/div/div/div[1]/div[1]/dx-card/div/dx-icon-title/div/p"));
			js = (JavascriptExecutor) driver;
			js.executeScript("arguments[0].scrollIntoView(true);", button3);
			click("SaveVehicle_XPATH");
			output("Vehicle Details saved successfully");
			test.log(Status.INFO, "Vehicle details is given successfully");
			//Thread.sleep(5000);
			// Coverage Details Tab
			if (data.get("Coverage Type").equalsIgnoreCase("Comprehensive")) {
//				WebElement button_cvg = driver.findElement(By.xpath("/html/body/dx-dijta-home-root/dx-vertical-layout/dx-layout/content/app-parent-qms/dx-compact-content/dx-layout-content/div/div/dx-qms-private-motor/div/dx-qms-nav-menu-layout/div[2]/div/div/app-private-motor-vehicle-details/form/div/dx-qms-nav-menu-content/div/div/div/div[1]/div[3]/div[4]/mat-accordion/mat-expansion-panel/div/div/dx-qms-risk-detail-sum-insure/div/div/div[2]/div[3]"));
//				js = (JavascriptExecutor) driver;
//				js.executeScript("arguments[0].scrollIntoView(true);", button_cvg);
				//Thread.sleep(5000);
				click("CovTpeDrpDon_XPATH");
				output("Coverage Type Drop Down Clicked successfully");
				// Enter("Comprehensive_XPATH",data.get("Coverage Type"));
				//Thread.sleep(5000);
				click("Compcoverclick_XPATH");
				output("Type of Coverage clicked successfully");
				//Thread.sleep(2000);
				Clear("VSI_XPATH");
				output("The Vehicle sum insured value cleared successfully");
				//Thread.sleep(5000);
				Enter("VSI_XPATH",data.get("VSI"));
				output("The Vehicle sum insured is given successfully");
				click("Agreed_XPATH");
				output("Agreed Value Check Box clicked successfully");
			} else if (data.get("Coverage Type").equalsIgnoreCase("Third Party")) {
				WebElement button_cvg = driver.findElement(By.xpath("//button[@id='cancel']"));
				js = (JavascriptExecutor) driver;
				js.executeScript("arguments[0].scrollIntoView(true);", button_cvg);
				//Thread.sleep(5000);
				click("CovTpeDrpDon_XPATH");
				output("Coverage Type Drop Down Clicked successfully");
				//Thread.sleep(5000);
				click("TPCoverclick_XPATH");
				output("Type of Coverage clicked successfully");
			}
			//Thread.sleep(2000);
			test.log(Status.INFO, "Coverage type details is given successfully");
//			String inceptionDate = print("InceptionDate_XPATH");
//			output("Inception date is: "+inceptionDate);
//			write(inceptionDate, m, 19, rowNum);
//			test.log(Status.INFO, "Inception date is: "+inceptionDate);
//			String expiryDate = print("ExpiryDate_XPATH");
//			output("Expiry date is: "+expiryDate);
//			write(expiryDate, m, 20, rowNum);
//			test.log(Status.INFO, "Expiry date is: "+expiryDate);
//			Thread.sleep(5000);
			Enter("PHID_XPATH", data.get("PH ID"));
			output("Policy Holder id Given successfully");
			Enter("nmeofID_XPATH", data.get("Name As Per ID"));
			output("Policy Holder name Given successfully");
			//Thread.sleep(5000);
			//		WebElement cancelButton = driver.findElement(By.xpath("//button[@id='cancel']"));
			//		js.executeScript("arguments[0].scrollIntoView(true);", cancelButton);
			//		Thread.sleep(10000);
			//		click("VldOnrISMBtn_XPATH");
			//		Thread.sleep(10000);
			//		output("Validate owner as per ISM button clicked successfullsy");
			click("Save2_XPATH");
			output("Save Quote Button clicked successfully");
			test.log(Status.INFO, "Save Quote is selected successfully");

			// Motor Additional Information
			// Additional Details of Proposer
			Thread.sleep(10000);
			//click("DrvExpDrpDwn_XPATH");
			lodclick("Drv2_XPATH");
			output("Driving Experience DropDown clicked successfully.");
			if (data.get("Driving Experience").equalsIgnoreCase("2 - 5 years")) {
				click("DrvExp2to5y_XPATH");
			} else if (data.get("Driving Experience").equalsIgnoreCase("6 - 10 years")) {
				click("DrvExp6to10y_XPATH");
			} else if (data.get("Driving Experience").equalsIgnoreCase(">10 years")) {
				click("DrvExp>10y_XPATH");
			} else if (data.get("Driving Experience").equalsIgnoreCase("Less than 2 years")) {
				click("DrvExp<2y_XPATH");
			}
			output("Driving Experience selected successfully");
			test.log(Status.INFO, "Driving Experience selected successfully");
			//Thread.sleep(5000);
//			Thread.sleep(15000);
//			click("HPToggle_XPATH");
//			output("Toggle Button is selected successfully");
//			Thread.sleep(15000);
//			click("HPDropdown_XPATH");
//			output("Dropdown is clicked successfully");
//			Enter("HPSel_XPATH", data.get("HP"));
//			output("Hire Purchase is selected successfully");
//			Thread.sleep(15000);
//			click("SelHP_XPATH");
//			output("The Hire Purchase is selected successfully");
			try{
				click("Yes_XPATH");
				output("Yes button is selected successfully");
			}
			catch(Exception e) {
				WebElement button_de = driver.findElement(By.xpath("/html/body/dx-dijta-home-root/dx-vertical-layout/dx-layout/content/app-parent-qms/dx-compact-content/dx-layout-content/div/div/dx-qms-private-motor/div/dx-qms-nav-menu-layout/div[2]/div/div/app-private-motor-coverage-details/div/dx-qms-nav-menu-content/div/div/div/div[1]/mat-accordion/mat-expansion-panel/div/div/app-motor-address/form/div/div[6]/div[1]"));
				js.executeScript("arguments[0].scrollIntoView(true);", button_de);
				click("State_XPATH");
				output("State Drop Down Clicked successfully");
				//Thread.sleep(1500);
				Enter("StateS_XPATH", data.get("Place Of Use2"));
				click("StateSel_XPATH");
				output("State Entered successfully");
				//Thread.sleep(2000);
				click("Post_XPATH");
				Enter("PostS_XPATH", data.get("PC").replace(".0", ""));
				click("PostSel_XPATH");
				output("Postal Code Entered successfully");
				//Thread.sleep(2000);
				click("Street_XPATH");
				//Thread.sleep(5000);
				//Enter("StreetS_XPATH", data.get("PC2"));
				//Thread.sleep(2000);
				click("Desa_XPATH");
				output("Street Entered successfully");
				//Thread.sleep(2000);
			}
			WebElement InvSD = driver.findElement(By.xpath(loc.getProperty("Inv_XPATH")));
			js.executeScript("arguments[0].scrollIntoView();", InvSD);
			// PIAM Analysis Code
			// Garage Type
			Thread.sleep(10000);
			click("GrgTpDR_XPATH");
			output("Garage Type Drop Down was Click successfully");
			//Thread.sleep(1500);
			Enter("GrgTypGivn_XPATH", data.get("Garage Type"));
			output("Garage Type Given successfully");
			//Thread.sleep(1000);
			click("GrgTypSect_XPATH");
			//Thread.sleep(1500);
			output("Garage Type Drop Down was selected successfully");
			// Anti Theft
			//Thread.sleep(3000);
			click("AntThfDR_XPATH");
			output("Anti Theft Drop Down was Click successfully");
			//Thread.sleep(1000);
			Enter("AntThtGivn_XPATH", data.get("Anti Theft"));
			output("Anti Theft Given successfully");
			//Thread.sleep(1500);
			click("AntThtSect_XPATH");
			//Thread.sleep(1000);
			output("Anti Theft Drop Down was selected successfully");
			// Safety Features
			//Thread.sleep(3000);
			click("SffeDr_XPATH");
			output("Safety Features Drop Down was Click successfully");
			//Thread.sleep(1500);
			Enter("SafFeuGivn_XPATH", data.get("Safety Features"));
			output("Safety Features Given successfully");
			//Thread.sleep(1000);
			click("SafFeuSect_XPATH");
			//Thread.sleep(5000);
			output("Safety Features Drop Down was selected successfully");
	
			test.log(Status.INFO, "PIAM Analysis Code Details Given successfully");


			// Generate Quote Letter
					click("GenQ_XPATH");
			//		click("DownQ_XPATH");
					click("Submit_XPATH");
					output("The Quote Letter is downloaded successfully");
			//		Thread.sleep(10000);
			//		WebElement messageElement = driver
			//				.findElement(By.xpath("//div[@aria-label='Quote Letter is Downloaded successfully!']"));
			//		if (messageElement.isDisplayed()) {
			//			test.pass(
			//					"<b>" + "<font color=" + "green>" + "Quote Letter is Downloaded successfully" + "</font>" + "</b>");
			//		} else {
			//			test.fail("<b>" + "<font color=" + "red>" + "Quote Letter is Downloaded successfully" + "</font>" + "</b>");
			//		}

			//Submit to TPM 
			//click("STPM_XPATH");
			//output("The Submit for TPM staff approval is clicked successfully");
			Thread.sleep(25000);
			click("ProceedHC2_XPATH");
			output("The Proceed policy cover is clicked successfully");
			Thread.sleep(25000);
			click("IssueCover_XPATH");
			output("The Issue Cover is clicked successfully");
			//Thread.sleep(45000);
//			click("PPS_XPATH");
//			click("PPSCB_XPATH");
//			click("PPSSubmit_XPATH");
//			output("The Policy Schedule is downloaded successfully");
			// Print
						Thread.sleep(25000);
						WebElement QuoteNum = Webprint("QuoteNumber_XPATH");
						String QuoteNumber = QuoteNum.getText();
						output("Quote Number is: " + QuoteNumber);
						write(QuoteNumber, m, 24, rowNum);

						WebElement QuoteSta = Webprint("QuoteStatus3_XPATH");
						String QuoteStatus = QuoteSta.getText();
						output("Quote Status is: " + QuoteStatus);
						write(QuoteStatus, m, 25, rowNum);
						WebElement PN = Webprint("PloicyNumber_XPATH");
						String policyNumber = PN.getText();
						output(policyNumber);
						write(policyNumber, m, 26, rowNum);
						write("Policy is Issued Successfully", m, 49, rowNum);
						test.pass("<b>" + "<font color=" + "green>" + "Policy is Issued successfully" + "</font>" + "</b>");

						WebElement VSI = Webprint("VehicleSumInsured_XPATH");
						String VehicleSumInsured = VSI.getText().replace(",", "");
						write(VehicleSumInsured, m, 27, rowNum);
						//Vehicle Sum Insured validation
						if (VehicleSumInsured.equalsIgnoreCase(data.get("Vehicle Sum Insured - Expected"))){
							write("Pass", m, 51, rowNum);
						} else {
							write("Fail", m, 51, rowNum);
						}

						WebElement Act = Webprint("ActPremium_XPATH");			
						String ActPremium = Act.getText();
						output("The Actual Premium is: "+ActPremium);
						write(ActPremium, m, 28, rowNum);
						//Actual Premium validation
						if (ActPremium.equalsIgnoreCase(data.get("Act Premium - Expected"))){
							write("Pass", m, 52, rowNum);
						} else {
							write("Fail", m, 52, rowNum);
						}

						WebElement Basic = Webprint("BasicPremium_XPATH");
						String BasicPremium = Basic.getText();
						output("The Basic Premium is: "+BasicPremium);
						write(BasicPremium, m, 29, rowNum);
						//Basic Premium validation
						if (BasicPremium.equalsIgnoreCase(data.get("Basic Premium - Expected"))){
							write("Pass", m, 53, rowNum);
						} else {
							write("Fail", m, 53, rowNum);
						}

						WebElement NCD2 = Webprint("NCD_XPATH");
						String NCD = NCD2.getText();
						output("The NCD value is: "+NCD);
						write(NCD, m, 30, rowNum);
						//NCD validation
						if (NCD.equalsIgnoreCase(data.get("NCD - Expected"))){
							write("Pass", m, 54, rowNum);
						} else {
							write("Fail", m, 54, rowNum);
						}
						Thread.sleep(5000);
						String PremiumafterNCD = driver.findElement(By.xpath(loc.getProperty("PremiumafterNCD_XPATH"))).getText();
						output("The Premium after NCD is: "+PremiumafterNCD);
						write(PremiumafterNCD, m, 31, rowNum);
						//Premium After NCD validation
						if (PremiumafterNCD.equalsIgnoreCase(data.get("Premium after NCD - Expected"))){
							write("Pass", m, 55, rowNum);
						} else {
							write("Fail", m, 55, rowNum);
						}

						WebElement EBT = Webprint("ExtraBenefitsTotal_XPATH");
						String ExtraBenefitsTotal = EBT.getText();
						output("The Extra Benefits Total value is: "+ExtraBenefitsTotal);
						write(ExtraBenefitsTotal, m, 32, rowNum);
						//Extra Benefits Total validation
						if (ExtraBenefitsTotal.equalsIgnoreCase(data.get("Extra Benefits Total - Expected"))){
							write("Pass", m, 56, rowNum);
						} else {
							write("Fail", m, 56, rowNum);
						}

						WebElement GP = Webprint("GrossPremium_XPATH");
						String GrossPremium = GP.getText();
						output("The Gross Premium value is: "+GrossPremium);
						write(GrossPremium, m, 33, rowNum);
						//GrossPremium validation
						if (GrossPremium.equalsIgnoreCase(data.get("Gross Premium - Expected"))){
							write("Pass", m, 57, rowNum);
						} else {
							write("Fail", m, 57, rowNum);
						}

						WebElement RP = Webprint("Rebate_XPATH");
						String Rebate = RP.getText();
						output("The Rebate value is: "+Rebate);
						write(Rebate, m, 34, rowNum);
						//Rebate validation
						if (Rebate.equalsIgnoreCase(data.get("Rebate - Expected"))){
							write("Pass", m, 58, rowNum);
						} else {
							write("Fail", m, 58, rowNum);
						}

						WebElement Tax = Webprint("SST_XPATH");
						String SST = Tax.getText();
						output("The SST value is: "+SST);
						write(SST, m, 35, rowNum);
						//SST validation
						if (SST.equalsIgnoreCase(data.get("SST - Expected"))){
							write("Pass", m, 59, rowNum);
						} else {
							write("Fail", m, 59, rowNum);
						}

						WebElement SD = Webprint("StampDuty_XPATH");
						String StampDuty = SD.getText();
						output("The Stamp Duty is: "+StampDuty);
						write(StampDuty, m, 36, rowNum);
						//StampDuty validation
						if (StampDuty.equalsIgnoreCase(data.get("Stamp Duty - Expected"))){
							write("Pass", m, 60, rowNum);
						} else {
							write("Fail", m, 60, rowNum);
						}

						WebElement TPP = Webprint("TotalPayablePremium_XPATH");
						//TotalPayablePremium is in  the range of Expected Premium
						String TotalPayablePremiumText = TPP.getText().trim();
						output("The Total Payable Premium is: " + TotalPayablePremiumText);
						write(TotalPayablePremiumText, m, 37, rowNum);


						// Retrieve Expected Premium and Actual/Expected Result
						String expectedPremiumText = data.get("Expected Premium") != null ? data.get("Expected Premium").trim() : "";
						String actualResult = data.get("Actual Result") != null ? data.get("Actual Result").trim() : "";
						String expectedResult = data.get("Expected Result") != null ? data.get("Expected Result").trim() : "";

						try {
							// Remove "MYR" and trim spaces
							TotalPayablePremiumText = TotalPayablePremiumText.replace("MYR ", "").trim();
							expectedPremiumText = expectedPremiumText.replace("MYR ", "").trim();

							// Convert to double
							double TotalPayablePremium = Double.parseDouble(TotalPayablePremiumText);
							double expectedPremium = Double.parseDouble(expectedPremiumText);

							// Calculate the range (±0.5% tolerance)
							double lowerBound = expectedPremium * 0.995;
							double upperBound = expectedPremium * 1.005;

							// Check if premium is within range
							boolean isPremiumMatched = (TotalPayablePremium >= lowerBound && TotalPayablePremium <= upperBound);
							boolean isResultMatched = actualResult.equals(expectedResult);

							if (isPremiumMatched && isResultMatched) {
								test.pass("<b><font color='green'>The premium is within the acceptable range</font></b>");
								write("Pass", m, 61, rowNum);
							} else {
								test.fail("<b><font color='red'>The Premium does not match within the range</font></b>");
								write("Fail", m, 61, rowNum);
							}
						} catch (NumberFormatException e) {
							test.fail("<b><font color='red'>Error: Unable to parse premium values</font></b>");
							output("Error parsing premium values: " + e.getMessage());
							write("Error", m, 61, rowNum);
						}
			// Checking the file size
			/*
			 * String downloadPath = "C:/Users/dell/Downloads/Quote - undefined.pdf"; File
			 * downloadedFile = new File(downloadPath); // Check if the file exists if
			 * (downloadedFile.exists()) { // Check the file size if
			 * (downloadedFile.length() == 0) { test.fail("<b>" + "<font color=" + "red>" +
			 * "Error: The downloaded policy kit file is 0 KB." + "</font>" + "</b>"); }
			 * else { test.pass("<b>" + "<font color=" + "green>" +
			 * "The downloaded policy kit file size is valid." + "</font>" + "</b>"); } }
			 * else { System.out.
			 * println("Error: The policy kit file was not found in the specified path."); }
			 */
			rowNum++;


			// Navigating to Home Screen
			Thread.sleep(15000);
			Baseclick("Home_XPATH");
			output("Home Button clicked successfully");
		}catch(Exception e) {
			System.out.println("Error in PCNV method: " + e.getMessage());
			//test.fail("Test case failed due to: " + e.getMessage());
			Thread.sleep(10000);
			rowNum++;
			WebElement element =  driver.findElement(By.xpath("//*[@id=\"dx-flex-menu-items\"]/a[1]/div"));
			element.click();
			click("leave_XPATH");
		}
		// ScreenRecord Stop
				Motor_PersonBP_ScreenRecord.stopRecord();
	}

	@AfterMethod(alwaysRun = true)
	public void tearDown(ITestResult result) {

		if (result.getStatus() == ITestResult.FAILURE) {
			String exceptionMessage = Arrays.toString(result.getThrowable().getStackTrace());
			test.fail("<details>" + "<summary><b><font color='red'>Exception Occurred: Click to see</font></b></summary>"
					+ exceptionMessage.replaceAll(",", "<br>") + "</details>\n");

			try {
				// Capture screenshot in Base64 format
				String base64Screenshot = TestUtil.captureScreenshot(driver);
				// Attach Base64 screenshot to the report
				test.fail("<b><font color='red'>Exception Occurred: Click to View Error Details</font></b>",
						MediaEntityBuilder.createScreenCaptureFromBase64String(base64Screenshot).build());
			} catch (IOException e) {
				e.printStackTrace();  // Log the error if screenshot capture fails
			}
			String failureLogg = "TEST CASE FAILED";
			Markup m = MarkupHelper.createLabel(failureLogg, ExtentColor.RED);
			test.log(Status.FAIL, m);
		}
		else if (result.getStatus() == ITestResult.SKIP) {

			String methodName = result.getMethod().getMethodName();

			String logText = "<b>" + "TEST CASE: - " + methodName.toUpperCase() + "  SKIPPED" + "</b>";

			Markup m = MarkupHelper.createLabel(logText, ExtentColor.YELLOW);
			test.skip(m);

		} else if (result.getStatus() == ITestResult.SUCCESS) {

			String methodName = result.getMethod().getMethodName();

			String logText = "<b>" + "TEST CASE: - " + methodName.toUpperCase() + "  PASSED" + "</b>";

			Markup m = MarkupHelper.createLabel(logText, ExtentColor.GREEN);
			test.pass(m);

		}
	}
	@AfterTest(alwaysRun = true)
	public void endReport() {
		extent.flush();
	}

}
