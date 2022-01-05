package basics;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.concurrent.TimeUnit;
import java.util.function.Function;

import org.apache.commons.io.FileUtils;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.impl.xb.xsdownload.DownloadedSchemaEntry;
import org.checkerframework.common.value.qual.StaticallyExecutable;
import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.Wait;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.gargoylesoftware.htmlunit.WebWindowAdapter;
import com.opencsv.CSVReader;
import com.opencsv.CSVWriter;
import com.opencsv.exceptions.CsvValidationException;

import ExcelFunctions.ConvertCSVfileToXLSXfile;
import seleniumpractice.delegation;

public class CSVReaderADMP1 
{
	static WebDriver driver;
	static boolean allMatched;

	//ConvertCSVfileToXLSXfile
	public static void ConvertCSVtoXLSX(String filereaderpath,String filewriterpath) throws CsvValidationException, IOException {
		FileReader filereader = new FileReader(filereaderpath);
		CSVReader csvReader = new CSVReader(filereader);
		String[] columns;

		XSSFWorkbook workbook=new XSSFWorkbook();
		XSSFSheet sheet=workbook.createSheet("sheet1");
		int rowcount=0;

		while ((columns = csvReader.readNext()) != null) 
		{
			XSSFRow row=sheet.createRow(rowcount++);
			int totalcsvcolumn=columns.length;
			for(int i=0;i<totalcsvcolumn;i++) 
			{
				XSSFCell cell=row.createCell(i);
				cell.setCellValue(columns[i]);
			}
		}
		csvReader.close();
		FileOutputStream fileOutputStream=new FileOutputStream(filewriterpath+".xlsx");
		workbook.write(fileOutputStream);
		workbook.close();
		fileOutputStream.close();
	}

	public static void checkingPDFdata(String starttime) throws IOException, CsvValidationException {
		FileInputStream inputStream=new FileInputStream(new File("E:\\selenium work\\New folder\\Output_of_ADMPsingleuser.xlsx"));
		XSSFWorkbook workbook=new XSSFWorkbook(inputStream);
		XSSFSheet sheet=workbook.getSheetAt(0);
		//int rowcount=sheet.getLastRowNum();
		XSSFRow row=sheet.getRow(0);
		int ColumnCount=row.getLastCellNum();
		XSSFCell cellfirstname=row.createCell(ColumnCount++);
		cellfirstname.setCellValue("FirstnamePDF");
		XSSFCell celllogonname=row.createCell(ColumnCount++);
		celllogonname.setCellValue("LogonnamePDF");
		XSSFCell cellcontainer=row.createCell(ColumnCount++);
		cellcontainer.setCellValue("EmailAddressPDF");
		XSSFCell cellfullname=row.createCell(ColumnCount++);
		cellfullname.setCellValue("FullnamePDF");


		File file = new File("C:\\Users\\com\\Downloads\\AuditReport.pdf");
		FileInputStream fileparse = new FileInputStream(file);
		PDDocument document = PDDocument.load(fileparse);
		System.out.println("Totalpages:"+document.getPages().getCount());
		PDFTextStripper pdfTextStripper = new PDFTextStripper();
		pdfTextStripper.setPageStart("Action Time : "+starttime);
		//pdfTextStripper.setPageEnd("Action Time : "+endtime);
		String doctext = pdfTextStripper.getText(document);
		//System.out.println(doctext);

		//CSV file reader
		String filereaderpath1="E:\\selenium work\\excel\\ADMPsingleuserdatacsv.csv";
		FileReader filereader = new FileReader(filereaderpath1);
		CSVReader csvReader = new CSVReader(filereader);
		String[] columns = csvReader.readNext();

		HashMap<String, Integer> map = new HashMap<>();
		int index=0;
		for(String columnName : columns)
		{
			map.put(columnName, index++);
		}
		int i=1;
		String[] lineInArray;
		while ((lineInArray = csvReader.readNext()) != null) 
		{
			String fname = lineInArray[map.get("First name")];
			String namelogon=lineInArray[map.get("Logon name")];
			String namefull=lineInArray[map.get("Full name")];
			String mailaddress=lineInArray[map.get("Email")];

			XSSFRow row2=sheet.getRow(i++);
			int TotalColumnCount=row2.getLastCellNum();
			XSSFCell cellfirstname1=row2.createCell(TotalColumnCount++);
			cellfirstname1.setCellValue(doctext.contains(fname));
			XSSFCell celllogonname1=row2.createCell(TotalColumnCount++);
			celllogonname1.setCellValue(doctext.contains(namelogon+"@admanagerplus.com"));
			XSSFCell cellmail=row2.createCell(TotalColumnCount++);
			cellmail.setCellValue(doctext.contains(mailaddress+"@admanagerplus.com"));
			XSSFCell cellfullname1=row2.createCell(TotalColumnCount++);
			cellfullname1.setCellValue(doctext.contains(namefull));
			System.out.println(doctext.contains(fname));

		}
		FileOutputStream fileOutputStream=new FileOutputStream("E:\\selenium work\\New folder\\Output_of_ADMPsingleuser.xlsx");
		workbook.write(fileOutputStream);
		workbook.close();
		inputStream.close();
		fileOutputStream.close();
		System.out.println("Done!");


		document.close();
		fileparse.close();
		System.out.println("Done!");

	}

	//screenshot
	public static void screenshot(String filename,String whichcase) throws IOException{
		String outputpath="E:\\selenium work\\output1\\Testcase"+whichcase+"\\"+filename+".jpg";
		File file=((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
		FileUtils.copyFile(file, new File(outputpath));
	}
	//date and time
	public static String datetime()
	{
		LocalDateTime dateTime=LocalDateTime.now();
		DateTimeFormatter dateFormat=DateTimeFormatter.ofPattern("dd-MM-yyyy_HH-mm-ss");
		String formattedDate=dateTime.format(dateFormat);
		return formattedDate;
	}
	public static String check_memberof_in_AR(String InputAttribute,String ARAttribute)
	{
		if(InputAttribute.equals(ARAttribute)) 
		{
			return "Matched";
		}
		else 
		{
			return "Not Matched"; 
		}
	}
	public static String checkingAuditReport(String InputAttribute,String ARAttribute) 
	{
		WebElement AuditReport=driver.findElement(By.xpath("//*[@class='mCSB_container']//*[contains(text(),'"+ARAttribute+"')]//following::div[1]"));
		JavascriptExecutor js= (JavascriptExecutor)driver;
		js.executeScript("arguments[0].scrollIntoView(true);", AuditReport);
		String string = AuditReport.getText();
		if(InputAttribute.equals(string)) 
		{
			return "Matched";
		}
		else 
		{
			allMatched = false;
			return "Not Matched"; 
		}
	}
	public static String checkOutputResult(String[] array)
	{
		for (String string : array) {
			if(string=="Not Matched")
			{
				return "Test Failed";
			}
		}
		return "Test Passed";
	}
	public static void DownloadPDFfile() throws InterruptedException
	{
		//Delegation
		driver.findElement(By.cssSelector("ul[id='top-menu'] li:nth-child(5)")).click();
		//screenshot("Delegation"+s1,path1);
		//Audit report
		driver.findElement(By.id("module_7003")).click();
		JavascriptExecutor js= (JavascriptExecutor)driver;
		WebDriverWait wait =new WebDriverWait(driver,10);
		Actions actions=new Actions(driver);
		driver.manage().timeouts().implicitlyWait(1,TimeUnit.MINUTES);
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@class='body-inner no-top-menu']//*[@id='selectedTech']//following-sibling::span")));
		WebElement selectdesk=driver.findElement(By.xpath("//div[@class='body-inner no-top-menu']//*[@id='selectedTech']//following-sibling::span"));
		js.executeScript("arguments[0].scrollIntoView(true);",selectdesk);
		js.executeScript("arguments[0].click();",selectdesk);
		//selectdesk.click();
		driver.findElement(By.cssSelector("tr:nth-child(2) [class='table-noborder bg_transparent']")).click();
		driver.findElement(By.xpath("//*[@id='onPopupOKBtn_10']//button[1]")).click();
		//select time
		WebElement time = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("[name='daterange']")));
		js.executeScript("arguments[0].click();",time);
		driver.findElement(By.cssSelector("[class='mCSB_container'] li:nth-child(6)")).click();
		//Go
		WebElement go = driver.findElement(By.id("generateReport"));
		js.executeScript("arguments[0].click();",go);
		//Export As
		driver.manage().timeouts().implicitlyWait(50,TimeUnit.SECONDS);
		List<WebElement> AR_datas = driver.findElements(By.cssSelector("[id='searchCol_auditReport']+tr td"));
		wait.until(ExpectedConditions.visibilityOfAllElements(AR_datas));
		driver.manage().timeouts().implicitlyWait(20,TimeUnit.SECONDS);
	
		wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("[class='dropdown tool-export']")));
		WebElement Exportas = driver.findElement(By.cssSelector("[data-original-title='Export as']"));
		actions.click(Exportas).build().perform();
		//actions.moveToElement(Exportas).build().perform();
		//Thread.sleep(3);
		//driver.manage().timeouts().implicitlyWait(10,TimeUnit.SECONDS);
		//actions.moveToElement(Exportas).click().build().perform();
		//Exportas.sendKeys(Keys.ENTER);
		//js.executeScript("arguments[0].click();",Exportas);
		//Exportas.click();
		//wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.cssSelector("[class*='dropdown tool-export'] li")));
		WebElement pdf=driver.findElement(By.cssSelector("[class*='dropdown tool-export'] li:nth-child(2)>a"));
		driver.manage().timeouts().implicitlyWait(1,TimeUnit.MINUTES);
		wait.until(ExpectedConditions.elementToBeClickable(pdf));
		js.executeScript("arguments[0].click();",pdf);
		System.out.println("Clicked PDf");
	}
	public static void main(String[] args) throws IOException, CsvValidationException, InterruptedException {
		// TODO
		String StartTime = null;
		String EndTime = null;
		String path1 = null;
		String filereaderpath = "E:\\selenium work\\excel\\ADMPsingleuserdatacsv1.csv";
		String filewriterpath = "E:\\selenium work\\excel";

		int i=1;

		System.setProperty("webdriver.chrome.driver","C:\\Program Files\\Java\\java selenium\\chromedriver\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(30,TimeUnit.SECONDS);//Implicit wait
		WebDriverWait wait =new WebDriverWait(driver,10);//Explicit wait
		Wait<WebDriver> fluentwait=new FluentWait<WebDriver>(driver)
				.withTimeout(Duration.ofSeconds(20))
				.pollingEvery(Duration.ofSeconds(3))
				.ignoring(NoSuchElementException.class);
		String defaultadmurl="http://demo.admanagerplus.com/#/mgmt";
		driver.navigate().to(defaultadmurl);
		//screenshot("ADManager",outputpath);
		JavascriptExecutor js= (JavascriptExecutor)driver;
		//Adminitrator
		Actions actions=new Actions(driver);
		WebElement clickadministator=driver.findElement(By.xpath("//a[contains(@onclick,'adminuser')]"));
		actions.moveToElement(clickadministator).click().build().perform();

		//CSV writer
		File file = new File("E:\\selenium work\\New folder\\Output_of_ADMPsingleuser.csv");
		FileWriter outputfile = new FileWriter(file);
		CSVWriter writer = new CSVWriter(outputfile);
		// adding header to csv
		String[] outputheader = { "LogonName","Password","Expected Result","office","Description","Full name","First name","Telephone number","Employee ID","Member of","Display name","Web page","logon name","Initails","Container name","Last name","Email ID","Test result"};
		writer.writeNext(outputheader);

		//CSV file reader
		FileReader filereader = new FileReader(filereaderpath);
		CSVReader csvReader = new CSVReader(filereader);
		//int totallines=csvReader.getMultilineLimit();
		//System.out.println(totallines);
		String[] columns = csvReader.readNext();

		HashMap<String, Integer> map = new HashMap<>();
		int index=0;

		for(String columnName : columns)
		{
			map.put(columnName, index++);
		}

		String[] lineInArray;

		while ((lineInArray = csvReader.readNext()) != null) 
		{
			String fname = lineInArray[map.get("First name")];
			String initials=lineInArray[map.get("Initials")];
			String lname=lineInArray[map.get("Last name")];
			String namelogon=lineInArray[map.get("Logon name")];
			String logon2000=lineInArray[map.get("Logonname prewindow")];
			String namefull=lineInArray[map.get("Full name")];
			String disname=lineInArray[map.get("Display name")];
			String employee=lineInArray[map.get("Employee ID")];
			String description=lineInArray[map.get("Description")];
			String office=lineInArray[map.get("Office")];
			String phonenum=lineInArray[map.get("Telephone number")];
			String mailaddress=lineInArray[map.get("Email")];
			String website=lineInArray[map.get("Webpage")];
			String container=lineInArray[map.get("Select container")];
			String password=lineInArray[map.get("Password")];
			String cpassword=lineInArray[map.get("Confirm password")];
			String memberof=lineInArray[map.get("MemberOf")];
			String[] array=memberof.split("\\\\");

			String s1=datetime();
			path1=i+"_"+s1;


			//management
			WebElement clickmanagement=fluentwait.until(new Function<WebDriver, WebElement>() {
				public WebElement apply(WebDriver driver) {
					return driver.findElement(By.xpath("//ul[@id='top-menu']/li[2]/a"));
				}
			});
			clickmanagement.click();
			screenshot("Click Managment"+s1,path1);
			//create single user
			WebElement clickSingle=driver.findElement(By.id("reportLink_6001"));
			wait.until(ExpectedConditions.visibilityOf(clickSingle));
			clickSingle.click();
			screenshot("create single user"+s1,path1);
			//firstname
			WebElement firstname=driver.findElement(By.cssSelector("#fieldinput2001>.form-control.form-control.input-md"));
			firstname.sendKeys(fname);
			screenshot("firstname"+s1,path1);
			//initals
			driver.findElement(By.id("input2002")).sendKeys(initials);
			screenshot("Initials"+s1, path1);
			//lastname
			driver.findElement(By.cssSelector("input#input2003")).sendKeys(lname);
			screenshot("lastname"+s1,path1);
			//logon name
			WebElement logonname=driver.findElement(By.cssSelector("input#input2004"));
			WebElement logonname1=wait.until(ExpectedConditions.visibilityOf(logonname));
			//TimeUnit.SECONDS.sleep(5);
			logonname1.clear();
			driver.manage().timeouts().implicitlyWait(10,TimeUnit.SECONDS);
			logonname1.sendKeys(namelogon);
			screenshot("Logonanme"+s1,path1);
			//logonname (pre-window 2000)
			WebElement logon=driver.findElement(By.cssSelector("input#input2005"));
			wait.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector("input#input2005")));
			//TimeUnit.SECONDS.sleep(5);
			logon.clear();
			driver.manage().timeouts().implicitlyWait(10,TimeUnit.SECONDS);
			logon.sendKeys(logon2000);
			screenshot("logonname(pre window)"+s1,path1);
			firstname.clear();
			firstname.sendKeys(fname);
			//full name
			WebElement fullname=driver.findElement(By.cssSelector("input#input2006"));
			driver.manage().timeouts().implicitlyWait(10,TimeUnit.SECONDS);
			fullname.clear();
			fullname.sendKeys(namefull);
			screenshot("Fullname"+s1,path1);
			//displayname
			WebElement displayname=driver.findElement(By.cssSelector("input#input2007"));
			driver.manage().timeouts().implicitlyWait(10,TimeUnit.SECONDS);
			displayname.clear();
			displayname.sendKeys(disname);
			screenshot("Displayname"+s1,path1);
			//Employee ID
			driver.findElement(By.id("input2008")).sendKeys(employee);
			screenshot("Employee ID"+s1,path1);
			//Description
			driver.findElement(By.id("input2009")).sendKeys(description);
			screenshot("Description"+s1,path1);
			//office
			driver.findElement(By.id("orgCaret_2010_span")).click();
			String officepath="//*[@id='OrgList_2010']//span[contains(text(),'"+office+"')]";
			driver.findElement(By.xpath(officepath)).click();
			screenshot("Office"+s1,path1);
			//telephone number
			driver.findElement(By.xpath("//li[@id=\"fieldinput2011\"]//input")).sendKeys(phonenum);
			screenshot("Telephonenumber"+s1,path1);
			//e-mail
			WebElement email=driver.findElement(By.cssSelector("input#input2012"));
			driver.manage().timeouts().implicitlyWait(10,TimeUnit.SECONDS);
			email.clear();
			email.sendKeys(mailaddress);
			driver.manage().timeouts().implicitlyWait(10,TimeUnit.SECONDS);
			//wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("inputDiv_2012"))).clear();
			//wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("userInput_2012"))).clear();
			driver.findElement(By.id("inputDiv_2012")).click();
			driver.manage().timeouts().implicitlyWait(10,TimeUnit.SECONDS);
			driver.findElement(By.id("userInput_2012")).click();
			screenshot("E-mail address"+s1,path1);
			//webpage
			driver.findElement(By.id("input2013")).sendKeys(website);
			screenshot("Webpage"+s1,path1);
			//select container
			WebElement clickcontainer=driver.findElement(By.id("selectLinkId_2014"));
			clickcontainer.click();
			//frame
			driver.switchTo().frame("userDetailsFrame");
			String containerpath="//li[contains(@id,'OU="+container+"')]//a//i[1]";
			//System.out.println(containerpath);
			WebElement containername=driver.findElement(By.xpath(containerpath));
			containername.click();
			screenshot("select container"+s1, path1);
			driver.switchTo().defaultContent();
			driver.findElement(By.id("popupButtonVal")).click();
			js.executeScript("window.scroll(0,0)"," ");
			//Account
			WebElement application=driver.findElement(By.xpath("//div[@title='Account'][@class='tabName']"));
			// js.executeScript("arguments[0].scrollIntoView(true);",application);
			application.click();
			screenshot("Account"+s1, path1);

			//type a password
			WebElement typepassword=driver.findElement(By.xpath("//*[@id='ownPassword_2015']//parent::div"));
			wait.until(ExpectedConditions.visibilityOf(typepassword));
			typepassword.click();
			driver.findElement(By.id("enterPassword_2015")).sendKeys(password);
			driver.findElement(By.id("confirmPassword_2015")).sendKeys(cpassword);
			screenshot("Password"+s1,path1);
			//Member of
			wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.xpath("//*[@id='importcsv2016']//preceding-sibling::div")));
			WebElement member=driver.findElement(By.xpath("//*[@id='importcsv2016']//preceding-sibling::div"));
			driver.manage().timeouts().implicitlyWait(30,TimeUnit.SECONDS);
			js.executeScript("window.scroll(0,100)"," ");
			member.click();

			screenshot("click importcsv2016"+s1,path1);
			WebElement addgroup=driver.findElement(By.xpath("//*[@id='addGroupBtn']"));
			wait.until(ExpectedConditions.visibilityOf(addgroup));
			addgroup.click();
			screenshot("click addGroup Button"+s1,path1);
			driver.findElement(By.xpath("//*[@class='search-base']//preceding::li[@class='search-field']")).click();
			screenshot("click search-base"+s1,path1);

			//enter the search value
			WebElement searchbox = driver.findElement(By.xpath("//*[@id='searchText_5701']"));
			searchbox.sendKeys(array[0]);
			searchbox.click();	
			screenshot("click searchText"+s1,path1);
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id='Result_5701']")));
			driver.findElement(By.xpath("//*[@id='Result_5701']"));
			//System.out.println(memberof);
			WebElement sElement = driver.findElement(By.xpath("//td[@title='"+array[1]+"']//preceding::td[1]"));
			js.executeScript("arguments[0].scrollIntoView(true);", sElement);
			sElement.click();
			wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='popupButtonVal']")));
			driver.findElement(By.xpath("//*[@id='popupButtonVal']")).click();
			screenshot("Clickok"+s1,path1);
			wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='setPrimaryBtn']//following::input[@name='finish']")));
			driver.manage().timeouts().implicitlyWait(10,TimeUnit.SECONDS);
			driver.findElement(By.xpath("//*[@id='setPrimaryBtn']//following::input[@name='finish']")).click();
			//create
			driver.manage().timeouts().implicitlyWait(10,TimeUnit.SECONDS);
			WebElement create = driver.findElement(By.xpath("//input[@name='save']"));
			create.click();


			String expectedresult = "Error Code 80072035 :Error in creating user, The server is unwilling to process the request.";
			WebElement result = driver.findElement(By.xpath("//div[@id='statusTable']//child::td[2]//child::span"));
			String firstoutput = result.getText();
			WebElement result2 = driver.findElement(By.xpath("//div[@id='statusTable']//child::td[2]//child::span[2]"));
			String secondoutput = result2.getText();
			String output = firstoutput+secondoutput;
			screenshot("Final output"+s1,path1);

			//Delegation
			driver.findElement(By.cssSelector("ul[id='top-menu'] li:nth-child(5)")).click();
			screenshot("Delegation"+s1,path1);
			//Audit report
			driver.findElement(By.id("module_7003")).click();
			screenshot("Audit report"+s1,path1);
			//search
			/*	WebElement searchelement = driver.findElement(By.id("colBasedSearch_auditReport"));
			js.executeScript("arguments[0].click();",searchelement);

			driver.manage().timeouts().implicitlyWait(50,TimeUnit.SECONDS);
			//wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.id("colBasedSearch_auditReport")));

			driver.findElement(By.id("searchValue_auditReport_OBJECT_NAME")).sendKeys(logon2000+Keys.ENTER);
			driver.manage().timeouts().implicitlyWait(50,TimeUnit.SECONDS);
			 */
			/////
			driver.manage().timeouts().implicitlyWait(50,TimeUnit.SECONDS);
			WebElement Auditreport = fluentwait.until(new Function<WebDriver,WebElement>() {
				public WebElement apply(WebDriver driver) {
					return driver.findElement(By.cssSelector("[class='floatThead-wrapper']"));
				}
			});
			
			//get time

			if(i==1) {
				WebElement starttime = driver.findElement(By.cssSelector("[id='searchCol_auditReport']+tr td:nth-child(6) div"));
				js.executeScript("arguments[0].scrollLeft = arguments[0].offsetWidth", starttime);
				StartTime = starttime.getText();
				System.out.println(StartTime);
			}
			//WebElement status=driver.findElement(By.cssSelector("[id='searchCol_auditReport']+tr td:nth-child(8) div"));
			//System.out.println(status.getText());

			//Details
			driver.manage().timeouts().implicitlyWait(30,TimeUnit.SECONDS);
			driver.manage().window().setSize(new Dimension(1440, 900));
			WebElement details = driver.findElement(By.cssSelector("[id='searchCol_auditReport']+tr td:nth-child(9) div"));
			wait.until(ExpectedConditions.elementToBeClickable(details));
			//js.executeScript("arguments[0].scrollIntoView(true);", details);
			js.executeScript("arguments[0].scrollLeft = arguments[0].offsetWidth", details);
			screenshot("Details"+s1,path1);
			actions.moveToElement(details).click().build().perform();
			//js.executeScript("arguments[0].click();",details);
			//details.click();
			driver.manage().window().maximize();
			String officeofARD = checkingAuditReport(office,"Office");
			String descriptionofARD = checkingAuditReport(description,"Description");
			String fullnameofARD = checkingAuditReport(namefull,"Full Name");
			String firstnameofARD = checkingAuditReport(fname,"First Name");
			String telephonenumofARD = checkingAuditReport(phonenum,"Telephone Number");
			screenshot("User Details1"+s1,path1);
			String employeeidofARD=checkingAuditReport(employee, "Employee ID");
			WebElement mElement=driver.findElement(By.xpath("//*[@class='mCSB_container']//*[contains(text(),'Member of')]//following::div[1]"));
			js.executeScript("arguments[0].scrollIntoView(true);", mElement);
			String memberofvalue = mElement.getText();
			String[] MoVsplit=memberofvalue.split(",");
			//System.out.println(MoVsplit[0]);
			String memberofARD=check_memberof_in_AR("[CN="+array[1],MoVsplit[0]);
			//System.out.println(memberofARD);
			screenshot("User Details2"+s1, path1);
			String displaynameofARD=checkingAuditReport(disname, "Display Name");
			String webpageofARD=checkingAuditReport(website, "Web Page");
			String logonnameofARD=checkingAuditReport(namelogon+"@admanagerplus.com","Logon Name");
			screenshot("User Details3"+s1,path1);
			//System.out.println(logon2000+"@admanagerplus.com");
			String initialsofARD=checkingAuditReport(initials,"Initials");
			String containernameofADR=checkingAuditReport("OU="+container+",DC=admanagerplus,DC=com","Container Name");
			String lastnameofADR=checkingAuditReport(lname,"Last Name");
			String emailaddressodARD=checkingAuditReport(mailaddress+"@admanagerplus.com","Email Address");
			screenshot("User Details4"+s1,path1);
			//String testcaseresult = allMatched ? "Test Passed" : "Test Failed";
			String[] strings= {officeofARD,descriptionofARD,fullnameofARD,firstnameofARD,telephonenumofARD,
					employeeidofARD,memberofARD,displaynameofARD,webpageofARD,logonnameofARD,initialsofARD,containernameofADR,
					lastnameofADR,emailaddressodARD};
			String testcaseresult=checkOutputResult(strings);

			String[] printdataOutput = {logon2000,password,output,officeofARD,descriptionofARD,fullnameofARD,firstnameofARD,telephonenumofARD,
					employeeidofARD,memberofARD,displaynameofARD,webpageofARD,logonnameofARD,initialsofARD,containernameofADR,
					lastnameofADR,emailaddressodARD,testcaseresult};
			driver.findElement(By.cssSelector("[class='btn btn-default btn-classic']:nth-of-type(1)")).click();
			writer.writeNext(printdataOutput);
			writer.flush();
			i++;
			if (expectedresult.equals(output))
			{
				System.out.println(output);
				System.out.println("Test passed");
			}
			else
			{
				System.out.println("Test failed");
			}

		}
		WebElement endtime=driver.findElement(By.cssSelector("[id='searchCol_auditReport']+tr td:nth-child(6) div"));
		js.executeScript("arguments[0].scrollLeft = arguments[0].offsetWidth", endtime);
		EndTime=endtime.getText();
		System.out.println(EndTime);
		DownloadPDFfile();
		Thread.sleep(60);

		//String pdfdata="matched";
		csvReader.close();
		writer.close();

		String CSVfilepath="E:\\selenium work\\New folder\\Output_of_ADMPsingleuser.csv";//filewriterpath+"Result_of_ADMPsingleuser.csv";
		String FileofXLSX="E:\\selenium work\\New folder\\Output_of_ADMPsingleuser";//filewriterpath+"Result_of_ADMPsingleuser";

		ConvertCSVtoXLSX(CSVfilepath,FileofXLSX);
		//sign out
		//driver.findElement(By.xpath("//*[text()='TalkBack']//following::li[1]")).click();
		//driver.findElement(By.xpath("//*[text()='Sign Out']")).click();
		
		//Excel color formatting
		FileInputStream inputStream=new FileInputStream(new File("E:\\selenium work\\New folder\\Output_of_ADMPsingleuser.xlsx"));
		XSSFWorkbook workbook=new XSSFWorkbook(inputStream);
		XSSFSheet sheet=workbook.getSheetAt(0);
		int rowcount=sheet.getLastRowNum();
		//System.out.println(rowcount);
		int j=1;
		while(j<rowcount+1) {
			XSSFRow row=sheet.getRow(j++);
			XSSFCellStyle style=workbook.createCellStyle();
			XSSFCell cell=row.getCell(17);
			String testcase=cell.getStringCellValue();
			if(testcase.equals("Test Passed")) 
			{
				style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
			}
			else
			{
				style.setFillForegroundColor(IndexedColors.RED.getIndex());
			}
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			cell.setCellStyle(style);
		}
		FileOutputStream fileOutputStream=new FileOutputStream("E:\\selenium work\\New folder\\Output_of_ADMPsingleuser.xlsx");
		workbook.write(fileOutputStream);
		System.out.println("test case color Done!");
		//checkingPDFdata(StartTime);
		
		FileInputStream inputStream1=new FileInputStream(new File("E:\\selenium work\\New folder\\Output_of_ADMPsingleuser.xlsx"));
		XSSFWorkbook workbook1=new XSSFWorkbook(inputStream1);
		XSSFSheet sheet1=workbook1.getSheetAt(0);
		//int rowcount=sheet.getLastRowNum();
		XSSFRow row=sheet1.getRow(0);
		int ColumnCount=row.getLastCellNum();
		XSSFCell cellfirstname=row.createCell(ColumnCount++);
		cellfirstname.setCellValue("FirstnamePDF");
		XSSFCell celllogonname=row.createCell(ColumnCount++);
		celllogonname.setCellValue("LogonnamePDF");
		XSSFCell cellcontainer=row.createCell(ColumnCount++);
		cellcontainer.setCellValue("EmailAddressPDF");
		XSSFCell cellfullname=row.createCell(ColumnCount++);
		cellfullname.setCellValue("FullnamePDF");
		Thread.sleep(60);
		File file1 = new File("C:\\Users\\com\\Downloads\\AuditReport.pdf");
		FileInputStream fileparse = new FileInputStream(file1);
		PDDocument document = PDDocument.load(fileparse);
		System.out.println(document.getPages().getCount());
		PDFTextStripper pdfTextStripper = new PDFTextStripper();
		pdfTextStripper.setPageStart("Action Time : "+StartTime);
		//pdfTextStripper.setPageEnd("Action Time : "+endtime);
		String doctext = pdfTextStripper.getText(document);
		//System.out.println(doctext);

		//CSV file reader
		String filereaderpath1="E:\\selenium work\\excel\\ADMPsingleuserdatacsv1.csv";
		FileReader filereader1 = new FileReader(filereaderpath1);
		CSVReader csvReader1 = new CSVReader(filereader1);
		String[] columns1 = csvReader1.readNext();

		String[] lineInArray1;
		int k=0;
		while ((lineInArray1 = csvReader1.readNext()) != null) 
		{
			String fname = lineInArray1[map.get("First name")];
			String namelogon=lineInArray1[map.get("Logon name")];
			String namefull=lineInArray1[map.get("Full name")];
			String mailaddress=lineInArray1[map.get("Email")];

			XSSFRow row2=sheet.getRow(k++);
			int TotalColumnCount=row2.getLastCellNum();
			XSSFCell cellfirstname1=row2.createCell(TotalColumnCount++);
			cellfirstname1.setCellValue(doctext.contains(fname));
			XSSFCell celllogonname1=row2.createCell(TotalColumnCount++);
			celllogonname1.setCellValue(doctext.contains(namelogon+"@admanagerplus.com"));
			XSSFCell cellmail=row2.createCell(TotalColumnCount++);
			cellmail.setCellValue(doctext.contains(mailaddress+"@admanagerplus.com"));
			XSSFCell cellfullname1=row2.createCell(TotalColumnCount++);
			cellfullname1.setCellValue(doctext.contains(namefull));
			System.out.println(doctext.contains(fname));

		}
		FileOutputStream fileOutputStream1=new FileOutputStream("E:\\selenium work\\New folder\\Output_of_ADMPsingleuser.xlsx");
		workbook.write(fileOutputStream1);
		workbook.close();
		inputStream.close();
		fileOutputStream.close();
		System.out.println("Done!");


		document.close();
		fileparse.close();
		System.out.println("Read PDF Done!");
		driver.quit();
		
	}

}
