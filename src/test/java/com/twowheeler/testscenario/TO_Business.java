package com.twowheeler.testscenario;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.Method;
import java.net.InetAddress;
import java.net.UnknownHostException;
import java.util.Arrays;
import java.util.Hashtable;

import javax.mail.MessagingException;
import javax.mail.internet.AddressException;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebElement;
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
import com.twowheeler.base.TestBaseTO;
import com.twowheeler.utilities.MonitoringMail;
import com.twowheeler.utilities.TestConfig;
import com.twowheeler.utilities.TestUtil;

import Utilities_UAT_ScreenRecord.Motor_PersonBP_ScreenRecord;

public class TO_Business extends TestBaseTO {

	public ExtentSparkReporter htmlReporter;
	public ExtentReports extent;
	public ExtentTest test;
	public JavascriptExecutor js;
	private int rowNum = 1;
	static String messageBody;
	static String testDataPath;
	static double finalSI;

	@BeforeTest(alwaysRun = true)
	public void setReport() throws IOException {
		//extent.setReportUsesManualConfiguration(true);

		htmlReporter = new ExtentSparkReporter("./reports/PCTO-UAT-Reports.html");

		htmlReporter.config().setEncoding("utf-8");
		htmlReporter.config().setDocumentTitle("Private Car UAT Automation Reports");
		htmlReporter.config().setReportName("Private Car UAT Automation Test Results");
		htmlReporter.loadXMLConfig(new File(".\\ExtentReportXMLs\\CustomeHeader.xml"));
		htmlReporter.config().setTheme(Theme.STANDARD);

		extent = new ExtentReports();
		extent.attachReporter(htmlReporter);

		extent.setSystemInfo("Automation Tester", "Ranadheer Reddy");
		extent.setSystemInfo("Organisation", "Serole Technologies");
	}

	@Test(dataProviderClass = TestUtil.class, dataProvider = "dp", alwaysRun = true)
	public void PCTO(Hashtable<String, String> data, Method m) throws Exception {
		// ScreenRecord Start
		Motor_PersonBP_ScreenRecord.startRecord("PersonBP");
		try {
		test = extent.createTest("Private Car Test");
		if (data.get("Runmode").equalsIgnoreCase("N")) {
			throw new SkipException("Skipping the test case " + "PCTO".toUpperCase() + " as the Run mode is NO");
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
		Baseclick("TOVeh_XPATH");
		output("TO Vehicle Selected Successfully");
		Baseclick("TOPCMC_XPATH");
		output("The TO for Private Car and Motorcycle");
		Baseclick("nxtbtn_XPATH");
		output("Next Button Clicked successfully");
		test.log(Status.INFO, "Type of Motor Car Details Given successfully");

		// Motor Coverage(s) Screen
		// Vehicle Search and Vehicle Details tab
		//Thread.sleep(5000);
		Enter("VeclregNo_XPATH", data.get("Vechcle Reg"));
		output("Vehicle Number Entered successfully");
		click("POUDrpDon_XPATH");
		output("Place Of Use Drop Down Clicked successfully");
		//Thread.sleep(1500);
		Enter("StateEntr_XPATH", data.get("Place Of Use"));
		output("State Entered successfully");
		//Thread.sleep(2000);
		click("Statesel_XPATH");
		output("Place Of Use selected successfully");
		//Thread.sleep(3500);
		click("VeclsrcBtn_XPATH");
		output("Vehicle search button clicked successfully");
		output("Vehicle Data retrieved from IRBM successfully");
		test.log(Status.INFO, "Vehicle Data retrieved from IRBM successfully");
		//Category
				click("Cat_XPATH");
				output("Category Drop Down Clicked successfully");
				//Thread.sleep(1500);
				Enter("CatEntr_XPATH", data.get("Category"));
				output("Category Entered successfully");
				//Thread.sleep(2000);
				click("CatSel_XPATH");
				output("Category selected successfully");
				//Body
						click("Body_XPATH");
						output("Body Drop Down Clicked successfully");
						//Thread.sleep(1500);
						Enter("BodyEntr_XPATH", data.get("Body Type"));
						output("Body Entered successfully");
						//Thread.sleep(2000);
						click("BodySel_XPATH");
						output("Body selected successfully");
						click("NewVecSve_XPATH");
						output("Vehicle Details saved successfully");
						test.log(Status.INFO, "Vehicle details is given successfully");
		// Coverage Details Tab
		if (data.get("Coverage Type").equalsIgnoreCase("Comprehensive")) {
			WebElement button = driver.findElement(By.xpath("//button[@id='cancel']"));
			js = (JavascriptExecutor) driver;
			js.executeScript("arguments[0].scrollIntoView(true);", button);
			//Thread.sleep(5000);
			click("CovTpeDrpDon_XPATH");
			//			click("CovNew_XPATH");
			output("Coverage Type Drop Down Clicked successfully");
			Enter("CovrvaleEntr_XPATH", data.get("Coverage Type"));
			output("Coverage Type Given successfully");
			//Thread.sleep(1000);
			click("CovrTypSel_XPATH");
			output("Coverage Type clicked successfully");
			//Thread.sleep(2000);
			//			String vehSIBefore = print("VehSI_XPATH").replace(",", "");
			//			System.out.println("The Vehicle Sum Insured is: "+vehSIBefore);
			//			double vehSIInt = Double.parseDouble(vehSIBefore);
			//			Thread.sleep(8000);
			//			double agreedValueInt = 10;
			//			double updatedSI = vehSIInt * agreedValueInt / 100;
			//			System.out.println(updatedSI);
			//			finalSI = vehSIInt + updatedSI;
			//			System.out.println(finalSI);
			String AgreedValue = read("PCTO", 6, rowNum);
			output("The agreed value is: "+AgreedValue);
			if (String.valueOf(AgreedValue).equals("0")) {
				output("Test case not related to Agreed Value");
			} else {
				//Thread.sleep(3000);
				click("Agreed_XPATH");
				output("Agreed Value Check Box clicked successfully");
				click("AdjAgreed_XPATH");
				output("Adjust Agreed Value Check Box clicked successfully");
				//Thread.sleep(1500);
				Clear("AgreedEntr_XPATH");
				output("Agreed Value cleared successfully");
				//Thread.sleep(2000);
				Enter("AgreedEntr_XPATH", data.get("Agreed Value").replace(".0", ""));
				output("Agreed Value extracted from excel and Given successfully");
				//				double agreedValueInt = 10;
				//				double updatedSI = vehSIInt * agreedValueInt / 100;
				//				System.out.println(updatedSI);
				//				finalSI = vehSIInt + updatedSI;
				//				System.out.println("The Vehicle sum insured after agreed value is: "+finalSI);
				//test.log(Status.INFO, "The Vehicle sum insured after agreed value is: "+finalSI);
			}

			
		} else if (data.get("Coverage Type").equalsIgnoreCase("Third Party")) {
			WebElement button = driver.findElement(By.xpath("//button[@id='cancel']"));
			js = (JavascriptExecutor) driver;
			js.executeScript("arguments[0].scrollIntoView(true);", button);
			//Thread.sleep(5000);
			click("CovTpeDrpDon_XPATH");
			//			click("CovNew_XPATH");
			output("Coverage Type Drop Down Clicked successfully");
			//Thread.sleep(5000);
			click("TPCoverclick_XPATH");
			//			click("TPNew_XPATH");
			output("Type of Coverage clicked successfully");
		}else if (data.get("Coverage Type").equalsIgnoreCase("Autobuddy")) {
			WebElement button = driver.findElement(By.xpath("//button[@id='cancel']"));
			js = (JavascriptExecutor) driver;
			js.executeScript("arguments[0].scrollIntoView(true);", button);
			//Thread.sleep(5000);
			click("AutbdtChkBox_XPATH");
			output("Type of Coverage clicked successfully");
			click("PlnDrClk_XPATH");
		output("Plan Type Drop Down Clicked successfully");
		Enter("PlnEnt_XPATH", data.get("Plan Type"));
		output("Plan Type Given successfully");
		//Thread.sleep(1000);
		click("PlnSel_XPATH");
		output("Plan Type clicked successfully");
		//Thread.sleep(2000);
		}
		test.log(Status.INFO, "Coverage Details Given successfully");
		//Thread.sleep(2000);
		String inceptionDate = print("InsDt_XPATH");
		output("Inception date is: "+inceptionDate);
		write(inceptionDate, m, 7, rowNum);
		test.log(Status.INFO, "The Inception Date is: "+inceptionDate);
		String expiryDate = print("ExpDt_XPATH");
		output("Expiry date is: "+expiryDate);
		write(expiryDate, m, 8, rowNum);
		test.log(Status.INFO, "The Expiry Date is: "+expiryDate);
		//Thread.sleep(5000);
//		click("Org_XPATH");
//		output("The Organisation Button is selected Successfully");
//		Thread.sleep(5000);
		Enter("PHID_XPATH", data.get("PH ID"));
		output("Policy Holder id Given successfully");
		Enter("nmeofID_XPATH", data.get("Name As Per ID"));
		output("Policy Holder name Given successfully");
		//Thread.sleep(5000);
		WebElement cancelButton = driver.findElement(By.xpath("//button[@id='cancel']"));
		js.executeScript("arguments[0].scrollIntoView(true);", cancelButton);
		//Thread.sleep(10000);
		//click("VldOnrISMBtn_XPATH");
		click("Validate_XPATH");
		//Thread.sleep(10000);
		output("Validate owner as per ISM button clicked successfully");
		// NCD Transfer Details
//				String NCDV = read("PCRV", 35, rowNum);
//				output(NCDV);
//							if (String.valueOf(NCDV).equals("NA")) {
//								output("Test case not related to NCD Transfer");
//							} else {
//								Thread.sleep(10000);
//								click("NCDTog_XPATH");
//								output("NCD Transfer Togile Clicked successfully");
//								Thread.sleep(3000);
//								Enter("PrvReg_XPATH", data.get("PVR"));
//								output("Prvevious Vechile Reg Entered successfully");
//								Thread.sleep(2500);
//								click("PrvSrhIcn_XPATH");
//								output("Search Icon Clicked successfully");
//								Thread.sleep(8000);
//								//click("CnfChkBox_XPATH");
//								//output("Confrim NCD Transfer? Check Box Clicked successfully");
//							}
		//click("SaveBtn_XPATH");
		//click("DevSaveBtn_XPATH");
		Thread.sleep(10000);
		click("SaveBtn2_XPATH");
		output("Save Button clicked successfully");
		test.log(Status.INFO, "Vehicle Details Given successfully");

		// Motor Additional Information
		// Additional Details of Proposer
		// Motor Additional Information
		// Additional Details of Proposer
		Thread.sleep(20000);
		//click("DrvExpDrpDwn_XPATH");
		click("Drv2_XPATH");
		output("Driving Experience DropDown clicked successfully");
		//Thread.sleep(1500);
		Enter("DrvExpGvn_XPATH", data.get("Driving Experience"));
		output("Driving Experience Given clicked successfully");
		//Thread.sleep(2000);
		click("DrvExpDrpDwnSel_XPATH");
		//Thread.sleep(5000);
		output("Driving Experience selected successfully");

//		Thread.sleep(15000);
//		click("HPToggle_XPATH");
//		output("Toggle Button is selected successfully");
//		Thread.sleep(15000);
//		click("HPDropdown_XPATH");
//		output("Dropdown is clicked successfully");
//		Enter("HPSel_XPATH", data.get("HP"));
//		output("Hire Purchase is selected successfully");
//		Thread.sleep(15000);
//		click("SelHP_XPATH");
//		output("The Hire Purchase is selected successfully");
				try{
					lodclick("Yes_XPATH");
					output("Yes button is selected successfully");
				}
				catch(Exception e) {
					WebElement button_de = driver.findElement(By.xpath("/html/body/dx-dijta-home-root/dx-vertical-layout/dx-layout/content/app-parent-qms/dx-compact-content/dx-layout-content/div/div/dx-qms-private-motor/div/dx-qms-nav-menu-layout/div[2]/div/div/app-private-motor-coverage-details/div/dx-qms-nav-menu-content/div/div/div/div[1]/mat-accordion/mat-expansion-panel/div/div/app-motor-address/form/div/div[6]/div[1]"));
					js.executeScript("arguments[0].scrollIntoView(true);", button_de);
					Addclick("State_XPATH");
					output("State Drop Down Clicked successfully");
					//Thread.sleep(1500);
					Enter("StateS_XPATH", data.get("Place Of Use2"));
					Addclick("StateSel_XPATH");
					output("State Entered successfully");
					//Thread.sleep(2000);
					Addclick("Post_XPATH");
					Enter("PostS_XPATH", data.get("PC").replace(".0", ""));
					Addclick("PostSel_XPATH");
					output("Postal Code Entered successfully");
					//Thread.sleep(2000);
					Addclick("Street_XPATH");
					//Thread.sleep(5000);
					//Enter("StreetS_XPATH", data.get("PC2"));
					//Thread.sleep(2000);
					Addclick("Desa_XPATH");
					output("Street Entered successfully");
					//Thread.sleep(2000);
				}
		// E-INvoce for Scroll Down
		//Thread.sleep(25000);
		WebElement InvSD = driver.findElement(By.xpath(loc.getProperty("Inv_XPATH")));
		js.executeScript("arguments[0].scrollIntoView();", InvSD);
		// Rebate
		//		String RebateValue = data.get("Rebate2");
		//		output(RebateValue);
		//		if (RebateValue.equals("NA")) {
		//			output("Test case not related to Rebate");
		//			test.log(Status.INFO, "Test case not related to Rebate");
		//		} else {
		//			Thread.sleep(10000);
		//			Enter("Rebt_XPATH", data.get("Rebate2"));
		//			output("Rebate extracted from excel and Given successfully");
		//			test.log(Status.INFO, "Test case related to Rebate so value fetched from Excel");
		//		}
		// PIAM Analysis Code
		// Garage Type
		Thread.sleep(25000);
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


		//Generate Quote Letter
				click("GenQ_XPATH");
		//		click("DownQ_XPATH");
				click("Submit_XPATH");
				output("The Quote Letter is generated successfully");

		// Policy Issuance
		Thread.sleep(65000);
		click("PolicyIssue_XPATH");
		output("Proceed to Policy Issuance Button Clicked successfully");
		// Issue Policy
		Thread.sleep(25000);
		//click("IssuePolicy_XPATH");
		click("IssuePolicy2_XPATH");
		output("Issue Policy Button Clicked successfully");
		Thread.sleep(45000);
		//click("PPS_XPATH");
		click("PPS2_XPATH");
		//click("PPSCB_XPATH");
		click("PPSSubmit_XPATH");
		output("The Policy Schedule is downloaded successfully");
//		Thread.sleep(35000);
//		click("Download_XPATH");
//		output("Download & e-mail Policy button clicked successfully");
//		click("genLetter_XPATH");
//		output("Download Policy Schedule checkbox clicked successfully");
//		click("submit_XPATH");
//		output("Submit button clicked successfully");
//		Thread.sleep(5000);
//		WebElement messageElement = driver
//				.findElement(By.xpath("//div[@aria-label='Policy Schedule is Downloaded successfully!']"));
//		if (messageElement.isDisplayed()) {
//			test.pass("<b>" + "<font color=" + "green>" + "Policy kit downloaded successfully" + "</font>" + "</b>");
//		} else {
//			test.fail("<b>" + "<font color=" + "red>" + "Policy kit not downloaded successfully" + "</font>" + "</b>");
//		}
		// Print
					Thread.sleep(40000);
					WebElement QuoteNum = Webprint("QuoteNumber_XPATH");
					String QuoteNumber = QuoteNum.getText();
					output("Quote Number is: " + QuoteNumber);
					write(QuoteNumber, m, 12, rowNum);

					WebElement QuoteSta = Webprint("QuoteStatus3_XPATH");
					String QuoteStatus = QuoteSta.getText();
					output("Quote Status is: " + QuoteStatus);
					write(QuoteStatus, m, 13, rowNum);
					WebElement PN = Webprint("PloicyNumber_XPATH");
					String policyNumber = PN.getText();
					output(policyNumber);
					write(policyNumber, m, 14, rowNum);
					write("Policy is Issued Successfully", m, 37, rowNum);
					test.pass("<b>" + "<font color=" + "green>" + "Policy is Issued successfully" + "</font>" + "</b>");

					WebElement VSI = Webprint("VehicleSumInsured_XPATH");
					String VehicleSumInsured = VSI.getText().replace(",", "");
					write(VehicleSumInsured, m, 15, rowNum);
					//Vehicle Sum Insured validation
					if (VehicleSumInsured.equalsIgnoreCase(data.get("Vehicle Sum Insured - Expected"))){
						write("Pass", m, 39, rowNum);
					} else {
						write("Fail", m, 39, rowNum);
					}

					WebElement Act = Webprint("ActPremium_XPATH");			
					String ActPremium = Act.getText();
					output("The Actual Premium is: "+ActPremium);
					write(ActPremium, m, 16, rowNum);
					//Actual Premium validation
					if (ActPremium.equalsIgnoreCase(data.get("Act Premium - Expected"))){
						write("Pass", m, 40, rowNum);
					} else {
						write("Fail", m, 40, rowNum);
					}

					WebElement Basic = Webprint("BasicPremium_XPATH");
					String BasicPremium = Basic.getText();
					output("The Basic Premium is: "+BasicPremium);
					write(BasicPremium, m, 17, rowNum);
					//Basic Premium validation
					if (BasicPremium.equalsIgnoreCase(data.get("Basic Premium - Expected"))){
						write("Pass", m, 41, rowNum);
					} else {
						write("Fail", m, 41, rowNum);
					}

					WebElement NCD2 = Webprint("NCD_XPATH");
					String NCD = NCD2.getText();
					output("The NCD value is: "+NCD);
					write(NCD, m, 18, rowNum);
					//NCD validation
					if (NCD.equalsIgnoreCase(data.get("NCD - Expected"))){
						write("Pass", m, 42, rowNum);
					} else {
						write("Fail", m, 42, rowNum);
					}
					Thread.sleep(5000);
					String PremiumafterNCD = driver.findElement(By.xpath(loc.getProperty("PremiumafterNCD_XPATH"))).getText();
					output("The Premium after NCD is: "+PremiumafterNCD);
					write(PremiumafterNCD, m, 19, rowNum);
					//Premium After NCD validation
					if (PremiumafterNCD.equalsIgnoreCase(data.get("Premium after NCD - Expected"))){
						write("Pass", m, 43, rowNum);
					} else {
						write("Fail", m, 43, rowNum);
					}

					WebElement EBT = Webprint("ExtraBenefitsTotal_XPATH");
					String ExtraBenefitsTotal = EBT.getText();
					output("The Extra Benefits Total value is: "+ExtraBenefitsTotal);
					write(ExtraBenefitsTotal, m, 20, rowNum);
					//Extra Benefits Total validation
					if (ExtraBenefitsTotal.equalsIgnoreCase(data.get("Extra Benefits Total - Expected"))){
						write("Pass", m, 44, rowNum);
					} else {
						write("Fail", m, 44, rowNum);
					}

					WebElement GP = Webprint("GrossPremium_XPATH");
					String GrossPremium = GP.getText();
					output("The Gross Premium value is: "+GrossPremium);
					write(GrossPremium, m, 21, rowNum);
					//GrossPremium validation
					if (GrossPremium.equalsIgnoreCase(data.get("Gross Premium - Expected"))){
						write("Pass", m, 45, rowNum);
					} else {
						write("Fail", m, 45, rowNum);
					}

					WebElement RP = Webprint("Rebate_XPATH");
					String Rebate = RP.getText();
					output("The Rebate value is: "+Rebate);
					write(Rebate, m, 22, rowNum);
					//Rebate validation
					if (Rebate.equalsIgnoreCase(data.get("Rebate - Expected"))){
						write("Pass", m, 46, rowNum);
					} else {
						write("Fail", m, 46, rowNum);
					}

					WebElement Tax = Webprint("SST_XPATH");
					String SST = Tax.getText();
					output("The SST value is: "+SST);
					write(SST, m, 23, rowNum);
					//SST validation
					if (SST.equalsIgnoreCase(data.get("SST - Expected"))){
						write("Pass", m, 47, rowNum);
					} else {
						write("Fail", m, 47, rowNum);
					}

					WebElement SD = Webprint("StampDuty_XPATH");
					String StampDuty = SD.getText();
					output("The Stamp Duty is: "+StampDuty);
					write(StampDuty, m, 24, rowNum);
					//StampDuty validation
					if (StampDuty.equalsIgnoreCase(data.get("Stamp Duty - Expected"))){
						write("Pass", m, 48, rowNum);
					} else {
						write("Fail", m, 48, rowNum);
					}

					WebElement TPP = Webprint("TotalPayablePremium_XPATH");
					//TotalPayablePremium is in  the range of Expected Premium
					String TotalPayablePremiumText = TPP.getText().trim();
					output("The Total Payable Premium is: " + TotalPayablePremiumText);
					write(TotalPayablePremiumText, m, 25, rowNum);


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
							write("Pass", m, 49, rowNum);
						} else {
							test.fail("<b><font color='red'>The Premium does not match within the range</font></b>");
							write("Fail", m, 49, rowNum);
						}
					} catch (NumberFormatException e) {
						test.fail("<b><font color='red'>Error: Unable to parse premium values</font></b>");
						output("Error parsing premium values: " + e.getMessage());
						write("Error", m, 49, rowNum);
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
		Thread.sleep(5000);
		Baseclick("Home_XPATH");
		output("Home Button clicked successfully");
	}catch(Exception e) {
		System.out.println("Error in PCRV method: " + e.getMessage());
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
