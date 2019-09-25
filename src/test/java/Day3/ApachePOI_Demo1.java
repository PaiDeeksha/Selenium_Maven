package Day3;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.markuputils.ExtentColor;
import com.aventstack.extentreports.markuputils.MarkupHelper;
import com.aventstack.extentreports.reporter.ExtentHtmlReporter;

public class ApachePOI_Demo1 {
  @Test
  public void f() throws IOException, InterruptedException {
	  File src= new File("C:\\Users\\training_b6B.01.16\\Desktop\\TestData.xlsx");
	  FileInputStream fis=new FileInputStream(src);
	  XSSFWorkbook WB= new XSSFWorkbook(fis);
	  XSSFSheet SH=WB.getSheetAt(0);
	  System.out.println("First row number "+SH.getFirstRowNum());
	  System.out.println("Last row number "+SH.getLastRowNum());
	  int RowCount=SH.getLastRowNum()-SH.getFirstRowNum();
	  System.out.println("The total row count is "+RowCount);
	  for(int i=1;i<=RowCount;i++)
	  {
	  System.out.println(SH.getRow(i).getCell(0).getStringCellValue()+"\t\t\t" + SH.getRow(i).getCell(1).getStringCellValue());
	  System.setProperty("webdriver.chrome.driver","C:\\Users\\training_b6B.01.16\\Desktop\\Browser Drivers_DP\\chromedriver.exe");
	  WebDriver driver= new ChromeDriver();
	  driver.get("http://10.232.237.143:443/TestMeApp/fetchcat.htm");
	  driver.findElement(By.linkText("SignIn")).click();
	  driver.findElement(By.name("userName")).sendKeys(SH.getRow(i).getCell(0).getStringCellValue());
      Thread.sleep(900);
	  driver.findElement(By.id("password")).sendKeys(SH.getRow(i).getCell(1).getStringCellValue());
	  Thread.sleep(900);
	  driver.findElement(By.name("Login")).click();
	  Thread.sleep(900);
	  driver.close();
	  
	  ExtentHtmlReporter reporter= new ExtentHtmlReporter("C:\\Users\\training_b6B.01.16\\Desktop\\NewExtentReport_DP.html");
	  ExtentReports extent=new ExtentReports();
	  extent.attachReporter(reporter);
	  ExtentTest logger=extent.createTest("DemoWebShop");
	  logger.addScreenCaptureFromPath("https://www.google.com/search?q=dog&rlz=1C1GCEB_enIN812IN814&source=lnms&tbm=isch&sa=X&ved=0ahUKEwjN79jL0NzkAhUEH48KHXFZAXUQ_AUIEigB&biw=1366&bih=657#imgrc=PhKxtUncoIln8M:");
      logger.log(Status.INFO, "Apache POI is used in this script");
      logger.log(Status.PASS, "Excel data reading is done successfully");
      logger.log(Status.FAIL, MarkupHelper.createLabel("This test case is failed",ExtentColor.BLACK));
      extent.flush();
      driver.close();
	  

 //Writing into excel file
	  /*XSSFRow header=SH.getRow(0);
	  XSSFCell header2=header.createCell(2);
	  header2.setCellValue("status");
	  SH.getRow(1).createCell(2).setCellValue("Pass");
	  SH.getRow(2).createCell(2).setCellValue("Fail");
	  FileOutputStream fos=new FileOutputStream(src);
	  WB.write(fos);*/
	  }
  }
}