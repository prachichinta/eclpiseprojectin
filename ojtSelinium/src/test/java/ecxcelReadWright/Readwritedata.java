package ecxcelReadWright;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Assert;
import org.testng.ITestResult;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class Readwritedata {
WebDriver driver=null;
int count=1;
	
	@Test(dataProvider="getLoginData")
	public void loginTest(String uname,String pass){
		String link="file:///D:/prachi%20clss/software/javabykiran-Selenium-Softwares/Offline%20Website/index.html";
		System.setProperty("webdriver.chrome.driver", "chromedriver.exe");
		driver=new ChromeDriver();
		driver.get(link);
		driver.findElement(By.xpath("/html/body/div/div[2]/form/div[1]/input")).sendKeys(uname);                //email
		driver.findElement(By.xpath("/html/body/div/div[2]/form/div[2]/input")).sendKeys(pass);              //password
		driver.findElement(By.xpath("/html/body/div/div[2]/form/div[3]/div/button")).click();                   //sign in	
	    Assert.assertEquals(driver.getTitle(), "JavaByKiran | DashBoard");
	}
	
	@AfterMethod
	public void writeResult(ITestResult result){
		if(result.isSuccess())
			ExcelReadWithPoi.writeDataInCell("logindata.xlsx", "login", count, 2, "Pass");
		else
			ExcelReadWithPoi.writeDataInCell("logindata.xlsx", "login", count, 2, "Fail");
	count++;
	}
	
	@DataProvider
	public String[][] getLoginData(){
		int rows=ExcelReadWithPoi.countRowFromExcel("logindata.xlsx", "login");
		int cols=ExcelReadWithPoi.countColmFromExcel("logindata.xlsx", "login",rows);
		String[][] data=new String[rows][cols];
		for(int i=1;i<=rows;i++){
            for(int j=0;j<cols;j++){
            	String value=ExcelReadWithPoi.getDataFromCell("logindata.xlsx", "login", i, j);
            data[i-1][j]=value;
            System.out.println(value);
            }
		}
		return data;
	}
}
