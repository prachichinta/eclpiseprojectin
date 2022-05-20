package ecxcelReadWright;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

public class ExcelReadWithPoi {
	static FileInputStream fis=null;
	static FileOutputStream fos=null;
	static Workbook wb=null;
	static Sheet sh=null;
	static Row r=null;
	WebDriver driver=null;
	@Test
	public void readAllDataFromExcel(){
		DataFormatter df=new DataFormatter();  //use this because vo kisi b type ka data string me covert karega
		try { 
			fis=new FileInputStream("logindata.xlsx");
			wb=WorkbookFactory.create(fis);
		} catch (Exception e) {
			e.printStackTrace();
		}
		sh= wb.getSheet("login");
		int rows=sh.getLastRowNum();
		for(int i=0;i<=rows;i++){
			int colm=sh.getRow(i).getLastCellNum();
			System.out.println("no of colm for row:"+i+" are "+colm);
			for(int j=0;j<colm;j++){
				Cell c=sh.getRow(i).getCell(j);
				System.out.println(df.formatCellValue(c)+" ");
			}
			System.out.println( );
		}
	}  
	

	public static String getDataFromCell(String filepath, String sheetName, int row, int colm) {
		DataFormatter df = new DataFormatter(); // use this because vo kisi b
												// type ka data string me covert
												// karega
		try {
			fis = new FileInputStream(filepath);
			wb = WorkbookFactory.create(fis);
		} catch (Exception e) {
			e.printStackTrace();
		}
		sh = wb.getSheet(sheetName);
		Cell c = sh.getRow(row).getCell(colm);
		return df.formatCellValue(c);

	}
	 
	@Test
	public void loginTest(){
		String link="file:///D:/prachi%20clss/software/javabykiran-Selenium-Softwares/Offline%20Website/index.html";
		System.setProperty("webdriver.chrome.driver", "chromedriver.exe");
		driver=new ChromeDriver();
		driver.get(link);
		String uname=getDataFromCell("loginData.xlsx","login",1,0);
		String pass=getDataFromCell("loginData.xlsx","login",1,1);;
		driver.findElement(By.xpath("//*[@id=\"email\"]")).sendKeys(uname);                //email
		driver.findElement(By.xpath("/html/body/div/div[2]/form/div[2]/input")).sendKeys(pass);              //password
		driver.findElement(By.xpath("/html/body/div/div[2]/form/div[3]/div/button")).click();                   //sign in
		System.out.println("sign in");	
	}
	
	public static void writeDataInCell(String filepath, String sheetName, int row, int colm, String value) {
		DataFormatter df = new DataFormatter(); // use this because vo kisi b
												// type ka data string me covert
												// karega
		try {
			fis = new FileInputStream(filepath);
			wb = WorkbookFactory.create(fis);
		} catch (Exception e) {
			e.printStackTrace();
		}
		if(wb.getSheet(sheetName)!=null)
		sh = wb.getSheet(sheetName);
		else
		sh=wb.createSheet(sheetName);
		
		if(sh.getRow(row)!=null)
			r=sh.getRow(row);
		else
			r=sh.createRow(row);
		if(r.getCell(colm)!=null)
			r.getCell(colm).setCellValue(value);
		else
		   r.createCell(colm).setCellValue(value);
		try{
			fos=new FileOutputStream(filepath);
			wb.write(fos);
			wb.close();
			fos.close();
		}catch(Exception e){
			e.printStackTrace();
		}	
	}
	
	
	public static int countRowFromExcel(String filepath, String sheetName){
		try{ 
			fis=new FileInputStream(filepath);
			wb=WorkbookFactory.create(fis);
		} catch (Exception e) {
			e.printStackTrace();
		}
		sh= wb.getSheet(sheetName);
		int rows=sh.getLastRowNum();
		return rows;
	}  
	public static int countColmFromExcel(String filepath, String sheetName, int row){
		try{ 
			fis=new FileInputStream(filepath);
			wb=WorkbookFactory.create(fis);
		} catch (Exception e) {
			e.printStackTrace();
		}
		sh= wb.getSheet(sheetName);
		int colm=sh.getRow(row).getLastCellNum();
		return colm;
	}  
	
	@Test
	public void writeExcel(){
		writeDataInCell("logindata.xlsx", "Sheet1", 6, 6, "1000");
	}

}
