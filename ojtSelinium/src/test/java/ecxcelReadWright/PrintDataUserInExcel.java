package ecxcelReadWright;

import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

public class PrintDataUserInExcel {
WebDriver driver=null;
int count=1;

	@Test
	public void printData(){
		String link="file:///D:/prachi%20clss/software/javabykiran-Selenium-Softwares/Offline%20Website/index.html";
		System.setProperty("webdriver.chrome.driver", "chromedriver.exe");
		WebDriver driver=new ChromeDriver();
		driver.get(link);
		driver.findElement(By.xpath("//*[@id=\"email\"]")).sendKeys("kiran@gmail.com");                //email
		driver.findElement(By.xpath("/html/body/div/div[2]/form/div[2]/input")).sendKeys("123456");              //password
		driver.findElement(By.xpath("/html/body/div/div[2]/form/div[3]/div/button")).submit();                   //sign in
		System.out.println("sign in");
		driver.findElement(By.xpath("//span[text()='Users']")).click();
		if(count==1){
			List<WebElement> headerRow= driver.findElements(By.xpath("//tr/th"));
			for(WebElement element : headerRow){
				ExcelReadWithPoi.writeDataInCell("logindata.xlsx", "sheet2", count-1, headerRow.indexOf(element), element.getText());	
			}
			count++;
		}
		while(count<=5){
			List<WebElement> Row= driver.findElements(By.xpath("//tr["+count+"]/td"));
			for(WebElement element : Row){
				ExcelReadWithPoi.writeDataInCell("logindata.xlsx", "sheet2", count-1, Row.indexOf(element), element.getText());	
			
			}
			
		count++;
	  }
	
   }
}
