import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject
import com.kms.katalon.core.checkpoint.Checkpoint as Checkpoint
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling as FailureHandling
import com.kms.katalon.core.testcase.TestCase as TestCase
import com.kms.katalon.core.testdata.TestData as TestData
import com.kms.katalon.core.testng.keyword.TestNGBuiltinKeywords as TestNGKW
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.common.WebUiCommonHelper
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import internal.GlobalVariable as GlobalVariable
import org.openqa.selenium.Keys as Keys
import java.time.format.DateTimeFormatter
import java.time.LocalDate
import com.kms.katalon.keyword.excel.ExcelKeywords as ExcelKeywords
//import alat.CharUtilitas as CharUtilitas



// for database
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
// for database


 
//Steps:
//1. Edit GetPrice.xlsx, semua isi kolom "Daily Done" dikosongin dan jumlah baris harus benar
//2. Set date on code, harus dalam 1 hari aja, karna hanya mengambil data harga hari ini saja
//3. hasilnya akan update data broksum & price per hari ini
//+ Gunakan Test suites, supaya jika TC failed, bs auto retry

// for ALL
String strTop1, strTop2, strTop3, strAllBrokAvg, strAccDist, strNetVal, strNetVol, strTanggal
String strBuyer1, strBuyerVol1, strBuyerAvg1, strSeller1, strSellerVol1, strSellerAvg1
String strBuyer2, strBuyerVol2, strBuyerAvg2, strSeller2, strSellerVol2, strSellerAvg2
String strBuyer3, strBuyerVol3, strBuyerAvg3, strSeller3, strSellerVol3, strSellerAvg3
String strOpenPrice, strLowPrice, strHighPrice, strClosePrice, strVol, strChange, strFreq, strFreqAnal
Integer intAvgRP
String strTemp
Double dblFreqAnal

LocalDate ldt = LocalDate.now()
String strNama
String strStartDate2 = "2024-01-10" // FR: disesuaikan
//String strStartDate2 = ldt.toString() // FR: disesuaikan

String strEndDate2 = "2024-01-10" // FR: disesuaikan
//String strEndDate2 = ldt.toString() // FR: disesuaikan
String[] partSplit
String strHari, strBulan, strTahun

StringBuilder builder = new StringBuilder();
LocalDate startDate = LocalDate.parse(strStartDate2);
LocalDate endDate = LocalDate.parse(strEndDate2);
LocalDate d = startDate;
// for ALL


// for GUI

WebUI.openBrowser('https://stockbit.com/login')
WebUI.maximizeWindow()
WebUI.setText(findTestObject('Stockbit/input_UserID'), 'fadli_romi@yahoo.com')
WebUI.setEncryptedText(findTestObject('Stockbit/input_Password'), 'qzE6rR3iJu4=')
WebUI.click(findTestObject('Stockbit/button_Login'))
//WebUI.waitForElementClickable(findTestObject('Stockbit/button_SkipAvatar'), 10)
//WebUI.click(findTestObject('Stockbit/button_SkipAvatar'))
WebUI.delay(3)
WebUI.navigateToUrl('https://stockbit.com/screener')



// for database
Connection connection = null;
Statement statement = null;
ResultSet resultSet = null;
Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
//String msAccDB = WebUI.concatenate([GlobalVariable.folderTestData, '\\Database1.accdb'] as String[])
String msAccDB = 'Database1.accdb'
String dbURL = "jdbc:ucanaccess://" + msAccDB;
String strSQL
connection = DriverManager.getConnection(dbURL);
statement = connection.createStatement();

// Step 1: Loading driver
try {
	Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
} catch (ClassNotFoundException cnfex) {
	System.out.println("Problem in loading or "
			+ "registering MS Access JDBC driver");
	cnfex.printStackTrace();
}
// Step 1: Loading driver
// for database

String excelFilePath = 'GetPrice.xlsx'
String sheetName = 'Sheet1'
workbook01 = ExcelKeywords.getWorkbook(excelFilePath)
sheet01 = ExcelKeywords.getExcelSheet(workbook01, sheetName)
int numTot = numTot = ExcelKeywords.getCellValueByIndex(sheet01, 0, 0)



while (d.isBefore(endDate) || d.equals(endDate)) {

	for (int rowIndex = 1; rowIndex < numTot + 1; rowIndex++) {
		strTemp = "NA"
		if (ExcelKeywords.getCellValueByIndex(sheet01, rowIndex, 4)!="DONE")
		{
			strNama = ExcelKeywords.getCellValueByIndex(sheet01, rowIndex, 0)
			WebUI.navigateToUrl('https://stockbit.com/#/symbol/' + strNama)
			WebUI.waitForElementClickable(findTestObject('Stockbit/menu_Bandarmology'), 10)
			WebUI.click(findTestObject('Stockbit/menu_Bandarmology'))
			strTanggal = d.format(DateTimeFormatter.ofPattern("MMM dd, yyyy"))
	
			
				
			strTanggal = d.format(DateTimeFormatter.ISO_DATE)
			System.out.println("strTanggal = " + strTanggal)
			System.out.println("strNama = " + strNama)
			System.out.println("Baris ke = " + rowIndex)
			
		
			/*
			partSplit = strTanggal.split("-");
			strHari = partSplit[2]
			//System.out.println("strHari: " + strHari)
			strBulan = partSplit[1]
			//System.out.println("strBulan: " + strBulan)
			strTahun = partSplit[0]
			//System.out.println("strTahun: " + strTahun)
			
			strTemp = strHari + "/" + strBulan + "/" + strTahun
			System.out.println("strTemp: " + strTemp)
			*/
			
		
			WebUI.click(findTestObject('Stockbit/input__BandarmologyOpenStartDate'))
			WebUI.sendKeys(findTestObject('Stockbit/input__BandarmologyOpenStartDate'), Keys.chord(Keys.CONTROL, 'a'))
			WebUI.sendKeys(findTestObject('Stockbit/input__BandarmologyOpenStartDate'), Keys.chord(Keys.BACK_SPACE))
			WebUI.setText(findTestObject('Stockbit/input__BandarmologyOpenStartDate'), strTanggal)
			WebUI.sendKeys(findTestObject('Stockbit/input__BandarmologyOpenStartDate'), Keys.chord(Keys.ENTER))
			//WebUI.delay(2)
			WebUI.click(findTestObject('Stockbit/input__BandarmologyOpenEndDate'))
			WebUI.sendKeys(findTestObject('Stockbit/input__BandarmologyOpenEndDate'), Keys.chord(Keys.CONTROL, 'a'))
			WebUI.sendKeys(findTestObject('Stockbit/input__BandarmologyOpenEndDate'), Keys.chord(Keys.BACK_SPACE))
			WebUI.setText(findTestObject('Stockbit/input__BandarmologyOpenEndDate'), strTanggal)
			WebUI.sendKeys(findTestObject('Stockbit/input__BandarmologyOpenEndDate'), Keys.chord(Keys.ENTER))
			WebUI.delay(2)
		
			
			
			//cek apakah sedang di-suspend
			//strTemp = WebUI.getText(findTestObject('Stockbit/Bandarmology_Top1Vol'))
			//System.out.println("Label suspend: " + strTemp)
			
			//Cek apakah ada transaksi
			strTemp = WebUI.getText(findTestObject('Stockbit/Bandarmology_AvgRP'))
			System.out.println("Bandarmology_AvgRP = " + strTemp)
			
			if (strTemp != "0") {
			//if (strTop1 != "-") {
				strTop1 = WebUI.getText(findTestObject('Stockbit/Bandarmology_Top1'))
				strTop3 = WebUI.getText(findTestObject('Stockbit/Bandarmology_Top3'))
				strTop5 = WebUI.getText(findTestObject('Stockbit/Bandarmology_Top5'))
				strAllBrokAvg = WebUI.getText(findTestObject('Stockbit/Bandarmology_TopAVG'))
				strTemp = WebUI.getText(findTestObject('Stockbit/Bandarmology_AvgRP'))
				if (strTemp.contains(",")) {
					strTemp = strTemp.replace(",", "")
				}
							
				
				//intAvgRP = WebUI.getText(findTestObject('Stockbit/Bandarmology_AvgRP')).toInteger()
				intAvgRP = strTemp.toInteger()
				strBuyer1 = WebUI.getText(findTestObject('Stockbit/Bandarmology_Buyer1'))
				strBuyerVol1 = WebUI.getText(findTestObject('Stockbit/Bandarmology_BuyerVol1'))
				strBuyerAvg1 = WebUI.getText(findTestObject('Stockbit/Bandarmology_BuyerAvg1'))
				if (strBuyerAvg1.contains(",")) {
					strBuyerAvg1 = strBuyerAvg1.replace(",", "")
				}
				//Hitung NetVol
				if (strBuyerVol1.contains(" B"))
				{
					strBuyerVol1 = strBuyerVol1.replace(" B", "000000")
				}else if (strBuyerVol1.contains(" M"))
				{
					//strBuyerVol1 = strBuyerVol1.replace(" M", "000")
					
					strBuyerVol1 = strBuyerVol1.replace(" M", "")
					strBuyerVol1 = strBuyerVol1.toDouble() * 1000000
				}
				else if (strBuyerVol1.contains(" K"))
				{
					//strBuyerVal1 = strBuyerVol1.replace(" K", "0")
					
					strBuyerVol1 = strBuyerVol1.replace(" K", "")
					strBuyerVol1 = strBuyerVol1.toDouble() * 1000
				}
				/*
				if (strBuyerVol1.contains(".")) {
					strBuyerVol1 = strBuyerVol1.replace(".", "")
				}
				*/
				
				//strBuyerVal1 = strBuyerVal1.toDouble() / strBuyerAvg1.toDouble()
				//System.out.println("strBuyerVal1 / NetVol !!!!!!!!! = " + strBuyerVal1)
				//Hitung NetVol
				
				
				strSeller1 = WebUI.getText(findTestObject('Stockbit/Bandarmology_Seller1'))
				strSellerVol1 = WebUI.getText(findTestObject('Stockbit/Bandarmology_SellerVol1'))
				strSellerAvg1 = WebUI.getText(findTestObject('Stockbit/Bandarmology_SellerAvg1'))
				if (strSellerAvg1.contains(",")) {
					strSellerAvg1 = strSellerAvg1.replace(",", "")
				}
				//Hitung NetVol
				if (strSellerVol1.contains(" B"))
				{
					strSellerVol1 = strSellerVol1.replace(" B", "000000")
				}
				else if (strSellerVol1.contains(" M"))
				{
					//strSellerVol1 = strSellerVol1.replace(" M", "000")
					
					strSellerVol1 = strSellerVol1.replace(" M", "")
					strSellerVol1 = strSellerVol1.toDouble() * 1000000
					
				}
				else if (strSellerVol1.contains(" K"))
				{
					//strSellerVol1 = strSellerVol1.replace(" K", "0")
					
					strSellerVol1 = strSellerVol1.replace(" K", "")
					strSellerVol1 = strSellerVol1.toDouble() * 1000
				}
				/*
				if (strSellerVol1.contains(".")) {
					strSellerVol1 = strSellerVol1.replace(".", "")
				}
				*/
				
				//strSellerVal1 = strSellerVal1.toDouble() / strSellerAvg1.toDouble()
				//System.out.println("strSellerVal1 / NetVol !!!!!!!!! = " + strSellerVal1)
				//Hitung NetVol
		
		
				if (WebUI.verifyElementNotPresent(findTestObject('Stockbit/Bandarmology_Buyer2'),1, FailureHandling.OPTIONAL)) {
					strBuyer2 = "-"
					strBuyerVol2 = "0"
					strBuyerAvg2 = "0"
					
				} else {
			
					strBuyer2 = WebUI.getText(findTestObject('Stockbit/Bandarmology_Buyer2'))
					strBuyerVol2 = WebUI.getText(findTestObject('Stockbit/Bandarmology_BuyerVol2'))
					strBuyerAvg2 = WebUI.getText(findTestObject('Stockbit/Bandarmology_BuyerAvg2'))
					if (strBuyer2 != "")
					{
						if (strBuyerAvg2.contains(",")) {
							strBuyerAvg2 = strBuyerAvg2.replace(",", "")
						}
						//Hitung NetVol
						if (strBuyerVol2.contains(" B"))
						{
							strBuyerVal2 = strBuyerVol2.replace(" B", "000000")
						}else if (strBuyerVol2.contains(" M"))
						{
							//strBuyerVal2 = strBuyerVol2.replace(" M", "000")
							
							strBuyerVol2 = strBuyerVol2.replace(" M", "")
							strBuyerVol2 = strBuyerVol2.toDouble() * 1000000
						}
						else if (strBuyerVol2.contains(" K"))
						{
							//strBuyerVal2 = strBuyerVol2.replace(" K", "0")
							
							strBuyerVol2 = strBuyerVol2.replace(" K", "")
							strBuyerVol2 = strBuyerVol2.toDouble() * 1000
						}
							
						if (strBuyerVol2.contains(".")) {
							strBuyerVal2 = strBuyerVol2.replace(".", "")
						}
						
						//strBuyerVal2 = (strBuyerVal2.toDouble() / strBuyerAvg2.toDouble()).toInteger()
						//System.out.println("strBuyerVal2 / NetVol !!!!!!!!! = " + strBuyerVal2)
						//Hitung NetVol
					}else {
						strBuyer2 = "-"
						strBuyerVol2 = "0"
						strBuyerAvg2 = "0"
					}
					
						
				}
				
		
				if (WebUI.verifyElementNotPresent(findTestObject('Stockbit/Bandarmology_Seller2'),1, FailureHandling.OPTIONAL)) {
					strSeller2 = "-"
					strSellerVol2 = "0"
					strSellerAvg2 = "0"
					
				} else {
			
					strSeller2 = WebUI.getText(findTestObject('Stockbit/Bandarmology_Seller2'))
					strSellerVol2 = WebUI.getText(findTestObject('Stockbit/Bandarmology_SellerVol2'))
					strSellerAvg2 = WebUI.getText(findTestObject('Stockbit/Bandarmology_SellerAvg2'))
					if (strSeller2 != "")
					{
						if (strSellerAvg2.contains(",")) {
							strSellerAvg2 = strSellerAvg2.replace(",", "")
						}
						//Hitung NetVol
						if (strSellerVol2.contains(" B"))
						{
							strSellerVol2 = strSellerVol2.replace(" B", "000000")
						}
						else if (strSellerVol2.contains(" M"))
						{
							//strSellerVol2 = strSellerVol2.replace(" M", "000")
							
							strSellerVol2 = strSellerVol2.replace(" M", "")
							strSellerVol2 = strSellerVol2.toDouble() * 1000000
						}
						else if (strSellerVol2.contains(" K"))
						{
							//strSellerVal2 = strSellerVol2.replace(" K", "0")
							
							strSellerVol2 = strSellerVol2.replace(" K", "")
							strSellerVol2 = strSellerVol2.toDouble() * 1000
						}
						/*
						if (strSellerVal2.contains(".")) {
							strSellerVal2 = strSellerVal2.replace(".", "")
						}
						*/
						//strSellerVal2 = (strSellerVal2.toDouble() / strSellerAvg2.toDouble()).toInteger()
						//System.out.println("strSellerVal2 / NetVol !!!!!!!!! = " + strSellerVal2)
						//Hitung NetVol
					}else {
						strSeller2 = "-"
						strSellerVol2 = "0"
						strSellerAvg2 = "0"
					}
					
						
				}
				
				
				if (WebUI.verifyElementNotPresent(findTestObject('Stockbit/Bandarmology_Buyer3'),1, FailureHandling.OPTIONAL)) {
					strBuyer3 = "-"
					strBuyerVol3 = "0"
					strBuyerAvg3 = "0"
					
				} else {
			
					strBuyer3 = WebUI.getText(findTestObject('Stockbit/Bandarmology_Buyer3'))
					strBuyerVol3 = WebUI.getText(findTestObject('Stockbit/Bandarmology_BuyerVol3'))
					strBuyerAvg3 = WebUI.getText(findTestObject('Stockbit/Bandarmology_BuyerAvg3'))
					if (strBuyer3 != "")
					{
						if (strBuyerAvg3.contains(",")) {
							strBuyerAvg3 = strBuyerAvg3.replace(",", "")
						}
						//Hitung NetVol
						if (strBuyerVol3.contains(" B"))
						{
							strBuyerVal3 = strBuyerVol3.replace(" B", "000000")
						}
						else if (strBuyerVol3.contains(" M"))
						{
							//strBuyerVol3 = strBuyerVol3.replace(" M", "000")
							
							strBuyerVol3 = strBuyerVol3.replace(" M", "")
							strBuyerVol3 = strBuyerVol3.toDouble() * 1000000
						}
						else if (strBuyerVol3.contains(" K"))
						{
							//strBuyerVol3 = strBuyerVol3.replace(" K", "0")
							
							strBuyerVol3 = strBuyerVol3.replace(" K", "")
							strBuyerVol3 = strBuyerVol3.toDouble() * 1000
						}
						/*
						if (strBuyerVal3.contains(".")) {
							strBuyerVal3 = strBuyerVal3.replace(".", "")
						}
						*/
						//strBuyerVal3 = (strBuyerVal3.toDouble() / strBuyerAvg3.toDouble()).toInteger()
						//System.out.println("strBuyerVal3 / NetVol !!!!!!!!! = " + strBuyerVal3)
						//Hitung NetVol
					} else {
						strBuyer3 = "-"
						strBuyerVol3 = "0"
						strBuyerAvg3 = "0"
					}
					
						
					
				}
				
				
				if (WebUI.verifyElementNotPresent(findTestObject('Stockbit/Bandarmology_Seller3'),1, FailureHandling.OPTIONAL)) {
					strSeller3 = "-"
					strSellerVol3 = "0"
					strSellerAvg3 = "0"
					
				} else {
			
					strSeller3 = WebUI.getText(findTestObject('Stockbit/Bandarmology_Seller3'))
					strSellerVol3 = WebUI.getText(findTestObject('Stockbit/Bandarmology_SellerVol3'))
					strSellerAvg3 = WebUI.getText(findTestObject('Stockbit/Bandarmology_SellerAvg3'))
					if (strSeller3 != "")
					{
						if (strSellerAvg3.contains(",")) {
							strSellerAvg3 = strSellerAvg3.replace(",", "")
						}
						//Hitung NetVol
						if (strSellerVol3.contains(" B"))
						{
							strSellerVol3 = strSellerVol3.replace(" B", "000000")
						}
						else if (strSellerVol3.contains(" M"))
						{
							//strSellerVol3 = strSellerVol3.replace(" M", "000")
							
							strSellerVol3 = strSellerVol3.replace(" M", "")
							strSellerVol3 = strSellerVol3.toDouble() * 1000000
							
						}
						else if (strSellerVol3.contains(" K"))
						{
							//strSellerVol3 = strSellerVol3.replace(" K", "0")
							
							strSellerVol3 = strSellerVol3.replace(" K", "")
							strSellerVol3 = strSellerVol3.toDouble() * 1000
						}
						/*
						if (strSellerVal3.contains(".")) {
							strSellerVal3 = strSellerVal3.replace(".", "")
						}
						*/
						//strSellerVal3 = (strSellerVal3.toDouble() / strSellerAvg3.toDouble()).toInteger()
						//System.out.println("strSellerVal3 / NetVol !!!!!!!!! = " + strSellerVal3)
						//Hitung NetVol
					}else {
						strSeller3 = "-"
						strSellerVol3 = "0"
						strSellerAvg3 = "0"
					}
					
						
				}
				
				
				strAccDist = WebUI.getText(findTestObject('Stockbit/Bandarmology_AccDist'))
				strNetVol =  WebUI.getText(findTestObject('Stockbit/Bandarmology_NetVol'))
				strNetVol = strNetVol.replace(",", "")
				strNetVal = WebUI.getText(findTestObject('Stockbit/Bandarmology_NetVal'))
				
				//CharUtilitas data = new CharUtilitas()
				WebUI.delay(3)
				strOpenPrice = WebUI.getText(findTestObject('Stockbit/Home_OpenPrice'))
				//strOpenPrice = data.RemoveChar(strOpenPrice, ",")
				strOpenPrice = strOpenPrice.replace(",", "")
				
				strLowPrice = WebUI.getText(findTestObject('Stockbit/Home_LowPrice'))
				//strLowPrice = data.RemoveChar(strLowPrice, ",")
				strLowPrice = strLowPrice.replace(",", "")
				
				strHighPrice = WebUI.getText(findTestObject('Stockbit/Home_HighPrice'))
				//strHighPrice = data.RemoveChar(strHighPrice, ",")
				strHighPrice = strHighPrice.replace(",", "")
				
				strClosePrice = WebUI.getText(findTestObject('Stockbit/Home_ClosePrice'))
				//strClosePrice = data.RemoveChar(strClosePrice, ",")
				strClosePrice = strClosePrice.replace(",", "")
				
				
				
				
				strVol = WebUI.getText(findTestObject('Stockbit/Home_Volume'))
				if (strVol.contains(" K")) {
					strVol = strVol.replace(" K", "000")
				}
				if (strVol.contains(",")) {
					strVol = strVol.replace(",", "")
				}
				
				
				
				
				strChange = WebUI.getText(findTestObject('Stockbit/Home_Change'))
				strChange = strChange.replace("%", "")
				//System.out.println("strChange = " + strChange)
				
								 
				/* FR: ini dipake
				strFreq = WebUI.getText(findTestObject('Stockbit/Home_Freq'))				
				if (strFreq.contains(",")) {
					strFreq = strFreq.replace(",", "")
				}
				if(strFreq == "NA" || strFreq == "")
				{
					strFreq = ""
				}else
				{
					
					//strFreqAnal =(strVol.toDouble() / (strFreq.toDouble() * strFreq.toDouble() * strFreq.toDouble()) * 100000000)
					dblFreqAnal =(strVol.toDouble() / (strFreq.toDouble() * strFreq.toDouble() * strFreq.toDouble()) * 100000000)					
					strFreqAnal = String.format("%.4f", dblFreqAnal)
					//System.out.println("Double Number: " + String.format("%.4f", strFreqAnal))
					System.out.println("Double strFreqAnal: " + strFreqAnal)
					
				}
				*/
				
				
				
				strTanggal = d.format(DateTimeFormatter.ofPattern("yyyy-MM-dd"))
				strTanggal = "#" + strTanggal + "#"
				strSQL = "select count(*) as JumBaris from TblBrokSum where Nama = '"+ strNama+"' and Tanggal = "+ strTanggal +""
				resultSet = statement.executeQuery(strSQL)
				resultSet.next();
				int intJumBaris = resultSet.getInt("JumBaris");
				if (intJumBaris > 0) {
					
					//update data
					System.out.println("Update data")
					
					try {
						
						strSQL = "update TblBrokSum set top1 = '"+ strTop1 +"', top2 = '"+ strTop3 +"', top3 = '"+ strTop5 +"', " +
						"AllBrokAvg = '"+ strAllBrokAvg +"', AccDist = '"+ strAccDist +"', AvgRP = "+ intAvgRP +", " +
						"Buyer1 = '"+ strBuyer1 +"', BuyerVal1 = '"+ strBuyerVol1 +"', BuyerAvg1 = "+ strBuyerAvg1 +", " +
						"Seller1 = '"+ strSeller1 +"', SellerVal1 = '"+ strSellerVol1 +"', SellerAvg1 = "+ strSellerAvg1 +", " +
						"Buyer2 = '"+ strBuyer2 +"', BuyerVal2 = '"+ strBuyerVol2 +"', BuyerAvg2 = "+ strBuyerAvg2 +", " +
						"Seller2 = '"+ strSeller2 +"', SellerVal2 = '"+ strSellerVol2 +"', SellerAvg2 = "+ strSellerAvg2 +", " +
						"Buyer3 = '"+ strBuyer3 +"', BuyerVal3 = '"+ strBuyerVol3 +"', BuyerAvg3 = "+ strBuyerAvg3 +", " +
						"Seller3 = '"+ strSeller3 +"', SellerVal3 = '"+ strSellerVol3 +"', SellerAvg3 = "+ strSellerAvg3 +", " +
						"NetVol = "+ strNetVol +", NetVal = '"+ strNetVal +"', OpenPrice = "+ strOpenPrice +", LowPrice = "+ strLowPrice +", " +
						"HighPrice = "+ strHighPrice +", ClosePrice = "+ strClosePrice +", Volume = '"+ strVol +"', Change = "+ strChange +", " + 
						"Frekuensi = "+ strFreq +", FrekAnal = "+ strFreqAnal +" " +
						"where Nama = '"+ strNama +"' and Tanggal = "+ strTanggal +""
		
						
						 System.out.println("strSQL = " + strSQL)
						 statement.executeUpdate(strSQL)
						 ExcelKeywords.setValueToCellByIndex(sheet01, rowIndex, 4, "DONE")
						 ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
				   
					} catch (SQLException sqlex) {
						sqlex.printStackTrace()
					}
					
					
				} else {
					
					try {
						System.out.println("Insert data")
						
				   
						strSQL = "insert into TblBrokSum (nama, tanggal, top1, top2, top3, AllBrokAvg, AccDist, AvgRP, " +
						"Buyer1, BuyerVal1, BuyerAvg1, Seller1, SellerVal1, SellerAvg1, Buyer2, BuyerVal2, BuyerAvg2, " +
						"Seller2, SellerVal2, SellerAvg2, Buyer3, BuyerVal3, BuyerAvg3, Seller3, SellerVal3, SellerAvg3, NetVol, NetVal, " +
						"OpenPrice, LowPrice, HighPrice, ClosePrice, Volume, Change, Frekuensi, FrekAnal) " +
						 "values ('"+ strNama +"', "+ strTanggal +", '"+ strTop1 +"', '"+ strTop3 +"', '"+ strTop5 +"', " +
						 "'"+ strAllBrokAvg +"', '"+ strAccDist +"', "+ intAvgRP +", '"+ strBuyer1 +"', '"+ strBuyerVol1 +"', "+ strBuyerAvg1 +", " +
						 "'"+ strSeller1 +"', '"+ strSellerVol1 +"', "+ strSellerAvg1 +", '"+ strBuyer2 +"', '"+ strBuyerVol2 +"', "+ strBuyerAvg2 +", " +
						 "'"+ strSeller2 +"', '"+ strSellerVol2 +"', "+ strSellerAvg2 +", '"+ strBuyer3 +"', '"+ strBuyerVol3 +"', "+ strBuyerAvg3 +", " +
						 "'"+ strSeller3 +"', '"+ strSellerVol3 +"', "+ strSellerAvg3 +", "+ strNetVol +", '"+ strNetVal +"', " +
						 ""+ strOpenPrice +", "+ strLowPrice +", "+ strHighPrice +", "+ strClosePrice +", '"+ strVol +"', "+ strChange +", "+ strFreq +", "+ strFreqAnal +");"
						
						 System.out.println("strSQL = " + strSQL)
						 statement.executeUpdate(strSQL)
						 ExcelKeywords.setValueToCellByIndex(sheet01, rowIndex, 4, "DONE")
						 ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
				   
					} catch (SQLException sqlex) {
						sqlex.printStackTrace()
					}
					
				}
				
		
			}
		
			
			
			WebUI.waitForElementClickable(findTestObject('Stockbit/menu_Bandarmology'), 10)
			WebUI.click(findTestObject('Stockbit/menu_Bandarmology'))
			
			
		} // End IF
		

		
	} // end for
	
		d = d.plusDays(1);
		

	
 } //end while


 
 if (null != connection)
	 {
	 
	 // cleanup resources, once after processing
	 if (resultSet != null)
	 {
		 resultSet.close()
	 }
	 
	 statement.close();
	 
	 // and then finally close connection
	 connection.close();
 }
 

 WebUI.click(findTestObject('Stockbit/img__Profile'))
 WebUI.click(findTestObject('Stockbit/button_Logout'))
 WebUI.closeBrowser()



