package com.twowheeler.utilities;

import com.twowheeler.base.TestBase;
import java.io.File;
import java.io.IOException;
import java.lang.reflect.Method;
import java.util.Base64;
import java.util.Date;
import java.util.Hashtable;
import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.testng.annotations.DataProvider;

public class TestUtil extends TestBase {
	
	public static String screenshotPath;
	public static String screenshotName;

	public static String captureScreenshot(WebDriver driver) throws IOException {

		// Capture screenshot as a file
        File scrFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
        
        // Convert screenshot to Base64
        String base64Screenshot = Base64.getEncoder().encodeToString(FileUtils.readFileToByteArray(scrFile));
 
        // Generate screenshot name
        Date d = new Date();
        screenshotName = d.toString().replace(":", "_").replace(" ", "_") + ".jpg";
 
        // Save screenshot file
        String screenshotPath = System.getProperty("user.dir") + "\\reports\\" + screenshotName;
        FileUtils.copyFile(scrFile, new File(screenshotPath));
        //TestUtil.screenshotName = screenshotPath;
 
        return "data:image/jpg;base64," + base64Screenshot;  // Returning Base64 image data
	}
	
	@DataProvider(name="dp")
	public Object[][] getData(Method m) {

		String sheetName = m.getName();
		int rows = excel.getRowCount(sheetName);
		int cols = excel.getColumnCount(sheetName);

		Object[][] data = new Object[rows - 1][1];
		
		Hashtable<String,String> table = null;

		for (int rowNum = 2; rowNum <= rows; rowNum++) { // 2

			table = new Hashtable<String,String>();
			
			for (int colNum = 0; colNum < cols; colNum++) {

				// data[0][0]
				table.put(excel.getCellData(sheetName, colNum, 1), excel.getCellData(sheetName, colNum, rowNum));
				data[rowNum - 2][0] = table;
			}

		}

		return data;

	}	
}
