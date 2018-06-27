package support.steps;


import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.Select;

import support.utils.OperationHelper;

public class StepDefinations {
	public static void main(String[] args) throws Throwable{
		
		OperationHelper support = new OperationHelper();
		support.openFileExcel("\\src\\test\\resources\\DataIn\\", "Elements.xlsx", "Login_Page");
		String[][] table = {
				{"Row_Index","Column_Index"},
				{"1","1"}
		};
		
		WebElement e = support.getElement(table);
		System.out.println(e);
		System.out.println("abc");
		

//		String URL = "file://sm-dev59/Temp/vu.hoang/SU3_Automation/PraticeEnvironments/dropdownlist.htm";
//		support.launch("chrome");
//		support.open(URL);
////		Select dll = new Select(support.getElement("demo-ddl-xp", "//select[@id='ddl']"));
////		support.selectDropdowList(dll, "2", "index");
//		support.selectDropdowList("demo-ddl-xp", "//select[@id='ddl']", "1", "index");
//		Thread.sleep(3000);
//		support.close();
//				
	}
}// end of class
