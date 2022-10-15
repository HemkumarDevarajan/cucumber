package org.baseclass;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class BaseClass {
	public static WebDriver driver;

	public static void launchBrowser() {
		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();
	}

	public static void maximize() {
		driver.manage().window().maximize();
	}

	public static void launchUrl(String value) {
		driver.get(value);
	}

	public static void elementClick(WebElement element) {
		element.click();
	}

	public static void send(WebElement element, String value) {
		element.sendKeys(value);
	}

	public static void jsClick(WebElement element) {
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("arguments[0].click()", element);

	}

	public static void jsSend(WebElement element, String setvalue) {
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("arguments[0].setAttribute('value','" + setvalue + "')", element);

	}

	public static void close() {
		driver.close();

	}

	public static String excelRead(String sheetname, String filename, int row, int cell) throws IOException {
		File f = new File("C:\\Users\\hemku\\eclipse-workspace\\Junit\\Text\\" + filename + ".xlsx");
		FileInputStream fis = new FileInputStream(f);
		Workbook w = new XSSFWorkbook(fis);
		Sheet s = w.getSheet(sheetname);
		for (int i = 0; i < s.getPhysicalNumberOfRows(); i++) {
			Row r = s.getRow(row);
			for (int j = 0; j < r.getPhysicalNumberOfCells(); j++) {
				Cell c = r.getCell(cell);
				int cellType = c.getCellType();
				String value;
				if (cellType == 1) {
					value = c.getStringCellValue();
					System.out.println(value);
				} else if (DateUtil.isCellDateFormatted(c)) {
					Date d = c.getDateCellValue();
					SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yy");
					value = sdf.format(d);
				} else {
					double numericCellValue = c.getNumericCellValue();
					long l = (long) numericCellValue;
					value = String.valueOf(l);
				}
				return value;

			}
		}
		return sheetname;
	}

	public static void excelwrite(String sheetname, int row, int cell, String value, String filename)
			throws IOException {
		File f = new File("C:\\Users\\hemku\\eclipse-workspace\\Junit\\Text\\" + filename + ".xlsx");
		Workbook w = new XSSFWorkbook();
		Sheet s = w.createSheet(sheetname);
		Row r = s.createRow(row);
		Cell c = r.createCell(cell);
		c.setCellValue(value);
		FileOutputStream fo = new FileOutputStream(f);
		w.write(fo);
	}

	public static void createRow(String sheetname, int row, int cell, String value, String filename)
			throws IOException {
		File f = new File("C:\\Users\\hemku\\eclipse-workspace\\Junit\\Text\\" + filename + ".xlsx");
		FileInputStream fi = new FileInputStream(f);
		Workbook w = new XSSFWorkbook(fi);
		Sheet s = w.getSheet(sheetname);
		Row r = s.createRow(row);
		Cell c = r.createCell(cell);
		c.setCellValue(value);
		FileOutputStream fos = new FileOutputStream(f);
		w.write(fos);
	}

	public static void createCell(String sheetname, int row, int cell, String value, String filename)
			throws IOException {
		File f = new File("C:\\Users\\hemku\\eclipse-workspace\\Junit\\Text\\" + filename + ".xlsx");
		FileInputStream fi = new FileInputStream(f);
		Workbook w = new XSSFWorkbook(fi);
		Sheet s = w.getSheet(sheetname);
		Row r = s.getRow(row);
		Cell c = r.createCell(cell);
		c.setCellValue(value);
		FileOutputStream fos = new FileOutputStream(f);
		w.write(fos);
	}

	public static void updateCell(String sheetname, int row, int cell, String oldvalue, String newvalue,
			String filename) throws IOException {
		File f = new File("C:\\Users\\hemku\\eclipse-workspace\\Junit\\Text\\" + filename + ".xlsx");
		FileInputStream fi = new FileInputStream(f);
		Workbook w = new XSSFWorkbook(fi);
		Sheet s = w.getSheet(sheetname);
		Row r = s.getRow(row);
		Cell c = r.getCell(cell);
		String cellValue = c.getStringCellValue();
		if (cellValue.equals(oldvalue)) {
			c.setCellValue(newvalue);
		}
		FileOutputStream fos = new FileOutputStream(f);
		w.write(fos);
	}
}
