package support.utils;

import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.Select;

public class OperationHelper {
	private FileInputStream file;
	private XSSFWorkbook workbook;
	private XSSFSheet sheet;
	private XSSFRow row;
	private XSSFCell cell;
	private String sheetNameTest = "";
	private WebDriver driver = null;
	
	
//================ MINH HONG ================
	
	public void openFileExcel(String filePath, String fileName, String sheetName) throws Throwable {
		String fileDir = System.getProperty("user.dir") + filePath + fileName;
		file = new FileInputStream(fileDir);
		workbook = new XSSFWorkbook(file);
		sheetNameTest = sheetName;
		sheet = workbook.getSheet(sheetNameTest);
	}

	public void changeToSheet(String sheetName) {
		sheetNameTest = sheetName;
		sheet = workbook.getSheet(sheetNameTest);
		System.out.println("Changed to sheet: " + sheetNameTest);
	}

	public String getValue_FromExcel(int rowIndex, int colIndex) {
		row = sheet.getRow(rowIndex);
		cell = row.getCell(colIndex);
		return cell.getStringCellValue().toString();
	}
	
	public WebElement getElement(String[][] table) {
		//1. From a datatable -> get row_eName, col_eName
		int row_table_default = 1; //Get row(1) of datatable
		int row_excel_index_eName = Integer.parseInt(table[row_table_default][0]);
		int col_excel_index_eName = Integer.parseInt(table[row_table_default][1]);
		
//		int row_excel_index = Integer.parseInt(table.get(row_table_default).get(0));
//		int col_excel_index = Integer.parseInt(table.get(row_table_default).get(1));
		
		String eName = getValue_FromExcel(row_excel_index_eName, col_excel_index_eName);
		String eLocator = getValue_FromExcel(row_excel_index_eName + 1, col_excel_index_eName);
		
		WebElement e = getElement(eName, eLocator);
		return e;
	}
	
	

	/**
	 * This function launch a specify browser
	 * AUTHOR: MINH.HONG
	 * CREATED DATE: JUN 26th 2018 (26/06/2018)
	 * @param browser			Browser
	 */
	public void launch(String browser) {
		browser.trim().toLowerCase();
		String rootPath = System.getProperty("user.dir");
		String chromeDriverPath = "\\src\\test\\references\\chromedriver.exe";
		String firefoxDriverPath = "\\src\\test\\references\\geckodriver.exe";

		if (browser == "chrome") {

			System.setProperty("webdriver.chrome.driver", rootPath + chromeDriverPath);
			driver = new ChromeDriver();
			driver.manage().window().maximize();

		} else if (browser == "firefox") {

			System.setProperty("webdriver.gecko.driver", rootPath + firefoxDriverPath);
			driver = new FirefoxDriver();
			driver.manage().window().maximize();

		} else {

			System.out.println("Please enter correct browser");

		}

	}
	
	

	/**
	 * This function opens an URL
	 * AUTHOR: MINH.HONG
	 * CREATED DATE: JUN 26th 2018 (26/06/2018)
	 * @param url			URL
	 */
	public void open(String url) {
		driver.get(url);
	}
	

	/**
	 * This function is close the browser
	 * AUTHOR: MINH.HONG
	 * CREATED DATE: JUN 26th 2018 (26/06/2018)
	 */
	public void close() {

		driver.close();
	}

	

		/**
		 * This function get an identifier
		 * AUTHOR: MINH.HONG
		 * CREATED DATE: JUN 26th 2018 (26/06/2018)
		 * @param eName			Name of element
		 * @param eLocator		Locator of element
		 */
	public By getIdentifier(String eName, String eLocator) {
		By identifier = null;
		if (eName.toLowerCase().endsWith("-xp")) {
			identifier = By.xpath(eLocator);
		} else if (eName.toLowerCase().endsWith("-id")) {
			identifier = By.id(eLocator);
		} else if (eName.toLowerCase().endsWith("-name")) {
			identifier = By.name(eLocator);
		} else if (eName.toLowerCase().endsWith("-css")) {
			identifier = By.cssSelector(eLocator);
		} else {
			System.out.println("Type isn't existed");
		}
		return identifier;
	}

	

	/**
	 * This function get an Element
	 * AUTHOR: MINH.HONG
	 * CREATED DATE: JUN 26th 2018 (26/06/2018)
	 * @param eName			Name of element
	 * @param eLocator		Locator of element
	 */
	public WebElement getElement(String eName, String eLocator) {
		WebElement e = null;
		e = driver.findElement(getIdentifier(eName, eLocator));
		return e;
	}
	

	/**
	 * This function selects a DropdowList
	 * AUTHOR: MINH.HONG
	 * CREATED DATE: JUN 26th 2018 (26/06/2018)
	 * @param eName			Name of element
	 * @param eLocator		Locator of element
	 * @param value			Value of element
	 * @param type			Type of DropdowList (index, value, text)
	 */
	public void selectDropdowList(String eName, String eLocator, String value, String type) {

		Select ddl = new Select(getElement(eName, eLocator));
		selectDropdowList(ddl, value, type);
		
	}
	
	/**
	 * This function selects a DropdowList
	 * AUTHOR: MINH.HONG
	 * CREATED DATE: JUN 26th 2018 (26/06/2018)
	 * @param ddl			DropdowList (Name and Locator of element)
	 * @param value			Value of element
	 * @param type			Type of DropdowList (index, value, text)
	 */
	public void selectDropdowList(Select ddl, String value, String type){
		
		if (type == "value") {
			ddl.selectByValue(value);
		} else if (type == "text") {
			ddl.selectByVisibleText(value);
		} else if (type == "index") {
			ddl.selectByIndex(Integer.parseInt(value));
		} else {
			System.out.println("Please input valid type");
		}
	}
	

}// end of class
