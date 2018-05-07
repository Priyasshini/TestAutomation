package com.test.sample.project.seleniumautomation;

import java.awt.Font;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;

import javax.imageio.ImageIO;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import ru.yandex.qatools.ashot.AShot;
import ru.yandex.qatools.ashot.Screenshot;
import ru.yandex.qatools.ashot.comparison.ImageDiff;
import ru.yandex.qatools.ashot.comparison.ImageDiffer;

//import org.apache.poi.ss.usermodel.CellStyle;
//import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;

public class ProductEnrichmentValidationWithMethods {
	static WebDriver driver = null;
	private static Map<String, String> resultsMap = new HashMap<String, String>();
	private static HashSet<String> keyset = new HashSet<String>();
	private static String variant = null;
	private static String size = null;

	
	@Test
	private static void start() throws Exception {

		System.setProperty("webdriver.chrome.driver",
				"C:\\Users\\apurva.c.sharma\\Downloads\\chromedriver_win32\\chromedriver.exe");

		driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.get("http://int-www2.hm.com/en_gb/index.html");
		if (driver.findElements(By.xpath("//button[@class='close icon-close-white js-close']")).size() != 0) {
			driver.findElement(By.xpath("//button[@class='close icon-close-white js-close']")).click();
		}

		openUrl();

	}

	private static void openUrl() {

		
		
		try {
			for (String url : keyset) {
				driver.get(url);
				Thread.sleep(2000);
				if ((driver.findElements(By.xpath("//div[@class='module product-description sticky-wrapper']"))
						.size() != 0)) {

					verifyProductImage(driver);
				} else {
					resultsMap.put(driver.getCurrentUrl(), "No Item Found error");
				}

			}
		} catch (InterruptedException e) {
			e.printStackTrace();

		}
	}

	private static void verifyProductImage(WebDriver driver) {
		BufferedImage expectedImage, expectedImage1, expectedImage2;
		try {
			SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy HH-mm-ss");
			Date date = new Date();
			expectedImage = ImageIO
					.read(new File(System.getProperty("user.dir") + "\\screenshots\\HMElementScreenshot.png"));

			expectedImage1 = ImageIO.read(new File(System.getProperty("user.dir") + "\\screenshots\\ExpImage.png"));

			expectedImage2 = ImageIO.read(new File(System.getProperty("user.dir") + "\\screenshots\\ExpImage2.png"));
			WebElement image = driver.findElement(By.xpath("//div[@class='product-detail-main-image-container']/img"));
			JavascriptExecutor jse = (JavascriptExecutor) driver;
			jse.executeScript("arguments[0].scrollIntoView();", image);

			Screenshot imageScreenshot = new AShot().takeScreenshot(driver, image);
			BufferedImage actualImage = imageScreenshot.getImage();

			ImageIO.write(imageScreenshot.getImage(), "PNG", new File(System.getProperty("user.dir") + "\\screenshots\\"
					+ dateFormat.format(date) + "_actualElementScreenshot.png"));

			ImageDiffer imgDiff = new ImageDiffer();
			ImageDiff diff = imgDiff.makeDiff(actualImage, expectedImage);

			ImageDiffer imgDiff1 = new ImageDiffer();
			ImageDiff diff1 = imgDiff1.makeDiff(actualImage, expectedImage1);

			ImageDiffer imgDiff2 = new ImageDiffer();
			ImageDiff diff2 = imgDiff2.makeDiff(actualImage, expectedImage2);
			// Enter if loop only if the product image is not H&M image
			if (diff.hasDiff()) {
				if (diff1.hasDiff()) {
					if (diff2.hasDiff()) {
						verifySizeIsEnabled(driver, "Product Image is displayed");
					} else {
						verifySizeIsEnabled(driver, "No Image is displayed");

					}
				} else {
					verifySizeIsEnabled(driver, "No Image is displayed");

				}
			} else {
				verifySizeIsEnabled(driver, "No Image is displaying");
			}

		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private static void verifySizeIsEnabled(WebDriver driver, String image) {
		try {
			if (driver
					.findElement(
							By.xpath("//div[@class='product-item-buttons']//div[contains(@class,'picker')]/button"))
					.isEnabled()) {

				driver.findElement(
						By.xpath("//div[@class='product-item-buttons']//div[contains(@class,'picker')]/button"))
						.click();

				// Getting all the sizes in a list
				List<WebElement> listSizes = driver.findElements(
						By.xpath("//div[@class='product-item-buttons']//div[contains(@class,'picker')]/ul/li"));
				for (int i = 1; i < listSizes.size(); i++) {
					// if (listSizes.get(i).isEnabled() &&
					// listSizes.get(i).isDisplayed()) {
					size = getText(driver,
							driver.findElement(By
									.xpath("//div[@class='product-item-buttons']//div[contains(@class,'picker')]/ul/li["
											+ (i + 1) + "]")));
					variant = driver
							.findElement(By.xpath("//ul[@class='picker-list is-inline']/li" + "[" + (i + 1) + "]"))
							.getAttribute("data-code");
					resultsMap.put(driver.getCurrentUrl() + " " + variant, image + " " + size);

				}
			} else {
				if (driver.findElement(By.xpath("//button[contains(@class,'button-big button-buy')]")).isEnabled()) {
					variant = "No Size";
					size = "Add to cart button is enabled";
					resultsMap.put(driver.getCurrentUrl() + " " + variant, image + " " + size);
				} else {
					variant = "No Size";
					size = "Add to cart button is disabled";
					resultsMap.put(driver.getCurrentUrl() + " " + variant, image + " " + size);

				}

			}
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	public static String getText(WebDriver driver, WebElement element) {
		return (String) ((JavascriptExecutor) driver).executeScript("return jQuery(arguments[0]).text();", element);
	}

	
	@BeforeTest
	public void readExcel() {
		try {
			File folder = new File("C:\\Users\\apurva.c.sharma\\validation\\seleniumautomation\\TestOutput");
			
			if (!folder.exists()) {
				folder.mkdirs();
				System.out.println("Created the output folder");
			}else {
				System.out.println("Folder already exists");
			}

			File f = new File("C:\\Users\\apurva.c.sharma\\validation\\seleniumautomation\\utils\\TestData.xlsx");
			FileInputStream fileIn = new FileInputStream(f);

			XSSFWorkbook workbook = new XSSFWorkbook(fileIn);
			// open sheet 0 which is first sheet of your worksheet
			XSSFSheet sheet = workbook.getSheetAt(0);

			for (Row row : sheet) {

				Cell c = row.getCell(0);
                 System.out.println(c);
				if (c == null) {
					// Nothing in the cell in this row, skip it
					System.out.println("Nothing in the cell in this row, skip it");
				} else {

					keyset.add(c.getStringCellValue());

				}

			}
			System.out.println(keyset);

		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	@AfterTest
	private static void printResults() {
		driver.quit();
		System.out.println("/**************** Results******************/");
		
		try {
			SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy HH-mm-ss");
			Date date = new Date();
			// FileOutputStream out = new FileOutputStream(new
			// File(System.getProperty("user.dir") + "\\utils"+"report.xlsx"));
			FileOutputStream out = new FileOutputStream(
					new File("C:\\Users\\apurva.c.sharma\\validation\\seleniumautomation\\TestOutput\\"
							+ dateFormat.format(date) + "report.xlsx"));
			// folders="C:\\Users\\apurva.c.sharma\\Desktop\\"+dateFormat.format(date)+"report.xlsx";

			XSSFWorkbook workbook2 = new XSSFWorkbook();
			XSSFSheet sheet2 = workbook2.createSheet("Test_Report");
			
			//
			/*XSSFCellStyle style = (XSSFCellStyle) workbook2.createCellStyle();
		    style.setFillBackgroundColor(IndexedColors.GREEN.getIndex());*/
//		    style.setFillForegroundColor(new XSSFColor(new java.awt.Color(141, 234, 249)));0, 128, 0
//		    style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
//		    XSSFFont font = workbook2.createFont();
//	            font.setColor(IndexedColors.RED.getIndex());
//	            style.setFont(font);
			
			
			//
//			CellStyle style = workbook.createCellStyle();
			
			XSSFCellStyle style = (XSSFCellStyle) workbook2.createCellStyle();
		    style.setFillForegroundColor(new XSSFColor(new java.awt.Color(128,0,0)));
		   style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			
		   
		   XSSFCellStyle stylesuccess = (XSSFCellStyle) workbook2.createCellStyle();
		   stylesuccess.setFillForegroundColor(new XSSFColor(new java.awt.Color(0,128,0)));
		   stylesuccess.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		   
			Row row = sheet2.createRow(0);
			int rownum = 1;
			Cell cell1 = row.createCell(0);
			
			cell1.setCellValue("URL");
//			cell1.setCellStyle(style);
			
			Cell cell2 = row.createCell(1);
			cell2.setCellValue("Message");
//			sheet2.autoSizeColumn(1);
			sheet2.setColumnWidth(1, 12000); 

//			cell2.setCellStyle(style);
			for (String key : resultsMap.keySet()) {

				row = sheet2.createRow(rownum++);
				Cell cellfirst = row.createCell(0);
				cellfirst.setCellValue((String) key);
				Cell cellsecond = row.createCell(1);
				cellsecond.setCellValue(resultsMap.get(key));
				if(resultsMap.get(key).contains("Product Image is displayed")) {
					System.out.println("successss");
					cellsecond.setCellStyle(stylesuccess);
				}
				if(resultsMap.get(key).contains("error") || resultsMap.get(key).contains("No Image is displayed")) {
					System.out.println("errorrrr ");
					cellsecond.setCellStyle(style);
				}
			}
			// for (String url : resultsMap.keySet()) {
			// String key = url.toString();
			// String value = resultsMap.get(url).toString();
			// System.out.println(key + " - " + value);
//			System.out.println("jSuccessfully created");

			// FileOutputStream os = new FileOutputStream("C:\\Report" + "/" + "Label_Data["
			// + dateFormat.format(date) + "].xlsx");
			// wb = gatherAllReportData(data, null);
			// wb.write(os);
			workbook2.write(out);

			out.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
//		for (String url : resultsMap.keySet()) {
//			String key = url.toString();
//			String value = resultsMap.get(url).toString();
//			System.out.println(key + " - " + value);
//		}

	}

}
