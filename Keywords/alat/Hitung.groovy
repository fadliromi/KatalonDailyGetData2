package alat

import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject
import com.kms.katalon.core.annotation.Keyword
import com.kms.katalon.core.checkpoint.Checkpoint
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling
import com.kms.katalon.core.testcase.TestCase
import com.kms.katalon.core.testdata.TestData
import com.kms.katalon.core.testobject.TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import java.time.LocalDate
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import alat.CekChangePrice as AccessCekChangePrice


import internal.GlobalVariable

public class Hitung {


	private String bndrValue, bndrAVG
	private Double dblTemp


	public Hitung(String BndrValue, String BndrAVG){
		bndrValue = BndrValue
		bndrAVG = BndrAVG
	}

	public String HitungTotNetVal () throws IOException {

		if (bndrAVG.contains(",")) {
			bndrAVG = bndrAVG.replace(",", "")
		}

		if (bndrValue.contains(" B")) {
			//bndrValue = bndrValue.replace(" B", "000000")
			bndrValue = bndrValue.replace(" B", "")
			dblTemp = bndrValue.toDouble() * 1000000000


			// dikali semilyar
			//bndrValue = bndrValue.toBigDecimal()
			//bndrValue = dblTemp * 1000000000
			bndrValue = String.format("%.2f",dblTemp)

		}else if (bndrValue.contains(" M")) {
			//bndrValue = bndrValue.replace(" M", "000")
			bndrValue = bndrValue.replace(" M", "")
			dblTemp = bndrValue.toDouble() * 1000000


			// dikali sejuta
			//bndrValue = bndrValue.toBigDecimal() * 1000000
			//bndrValue = dblTemp * 1000000
			bndrValue = String.format("%.2f",dblTemp)

		}else if (bndrValue.contains(" K")) {
			//bndrValue = bndrValue.replace(" K", "")
			bndrValue = bndrValue.replace(" K", "")
			dblTemp = bndrValue.toDouble() * 1000

			// dikali seribu
			//bndrValue = bndrValue.toBigDecimal() * 1000
			//bndrValue = dblTemp * 1000
			bndrValue = String.format("%.2f",dblTemp)

		}
		/*
		 if (bndrValue.contains(".")) {
		 bndrValue = bndrValue.replace(".", "")
		 }
		 */
		//bndrValue = (bndrValue.toDouble() / bndrAVG.toDouble()).toInteger()
		//System.out.println("not contains(-), bndrValue = " + bndrValue)
		//bndrValue = bndrValue.toBigInteger()
		return bndrValue
	}
}



public class Split {

	private String tanggal, strHari, strBulan, strTahun
	private String[] partSplit

	public Split(String Tanggal) {
		tanggal = Tanggal
	}

	public String SplitTanggal () throws IOException {

		partSplit = tanggal.split("-");
		strHari = partSplit[2]
		//System.out.println("strHari: " + strHari)
		strBulan = partSplit[1]
		//System.out.println("strBulan: " + strBulan)
		strTahun = partSplit[0]
		//System.out.println("strTahun: " + strTahun)
		tanggal = strHari + "/" + strBulan + "/" + strTahun
		//System.out.println("strTempTanggal: " + tanggal)

		return tanggal

	}
}


public class HitungVolMA {

	private String strTanggal
	int intWorkDay, intMA

	public HitungVolMA(String str, int intTemp) {
		strTanggal = str
		intMA = intTemp
	}

	public String CariTglSebelumnya () throws IOException {
		//mencari tanggal (weekday), n hari sebelumnya

		LocalDate result = LocalDate.parse(strTanggal)

		//while (intWorkDay < 20)
		while (intWorkDay < intMA)
		{
			result = result.minusDays(1);
			if (result.getDayOfWeek().toString() != "SATURDAY" && result.getDayOfWeek().toString() != "SUNDAY")
			{
				System.out.println("ini weekday")
				System.out.println("Tanggal: " + result.toString() + ", "+ result.getDayOfWeek().toString())

				++intWorkDay
			}

		}

		return result.toString()
	}
}

public class CharUtilitas {


	public CharUtilitas()
	{
		//strCharacter = chr
		//strInput = strInput
	}

	public String RemoveChar (String strInput, String strCharacter) throws IOException
	{

		if (strInput.contains(strCharacter))
		{
			strInput = strInput.replace(strCharacter, "")
		}

		return strInput

	}

}

public class AccessDatabase
{
	Connection connection = null
	Statement statement = null

	public AccessDatabase()
	{

	}

	public OpenCon () throws IOException
	{
		Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
		String msAccDB = WebUI.concatenate([GlobalVariable.folderTestData, '\\Database1.accdb'] as String[])
		String dbURL = "jdbc:ucanaccess://" + msAccDB;

		connection = DriverManager.getConnection(dbURL);
		//statement = connection.createStatement();

		// Step 1: Loading driver
		try {
			Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
		} catch (ClassNotFoundException cnfex) {
			System.out.println("Problem in loading or "
					+ "registering MS Access JDBC driver");
			cnfex.printStackTrace();
		}
		// Step 1: Loading driver

		return connection

	}

}

public class AccessDatabase2
{
	Connection connection = null
	Statement statement = null

	public AccessDatabase2()
	{

	}

	public OpenCon () throws IOException
	{
		Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
		String msAccDB = WebUI.concatenate([GlobalVariable.folderTestData, '\\StudiKasus\\Marked.accdb'] as String[])
		String dbURL = "jdbc:ucanaccess://" + msAccDB;

		connection = DriverManager.getConnection(dbURL);
		//statement = connection.createStatement();

		// Step 1: Loading driver
		try {
			Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
		} catch (ClassNotFoundException cnfex) {
			System.out.println("Problem in loading or "
					+ "registering MS Access JDBC driver");
			cnfex.printStackTrace();
		}
		// Step 1: Loading driver

		return connection

	}

}

public class Scalping
{
	//private String strNmShm1, strNmShm2, strNmShm3, strNmShm4
	private String[][] strNmShm = new String[2][2]
	//private Integer intJmlItem
	String strAVGPrice

	public Scalping()
	{

	}

	public String getAVGPrice (String[] strNmShm) throws IOException {

		WebUI.navigateToUrl('https://stockbit.com/#/symbol/' + strNmShm)
		WebUI.delay(2)
		strAVGPrice = WebUI.getText(findTestObject('Stockbit/Home_AVGPrice'))


		return strAVGPrice
	}

	public String[][] getAVGPrice (String[][] strNmShm, Integer intJmlItem) throws IOException {

		//System.out.println("strNmShm[0,0]= " + strNmShm[0][0])

		for (int rowIndex = 0; rowIndex < intJmlItem; rowIndex++)
		{
			WebUI.navigateToUrl('https://stockbit.com/#/symbol/' + strNmShm[rowIndex][0])
			strNmShm[rowIndex][1] = WebUI.getText(findTestObject('Stockbit/Home_AVGPrice'))
		}

		//strNmShm[0][1] = "111"
		//strNmShm[1][1] = "222"

		/*
		 WebUI.navigateToUrl('https://stockbit.com/#/symbol/' + strNmShm[0][0])
		 strNmShm[0][1] = WebUI.getText(findTestObject('Stockbit/Home_AVGPrice'))		
		 WebUI.navigateToUrl('https://stockbit.com/#/symbol/' + strNmShm[1][0])
		 strNmShm[1][1] = WebUI.getText(findTestObject('Stockbit/Home_AVGPrice'))
		 */


		return strNmShm
	}

}


public class CheckAkumSama {

	private String strTanggal, strSQL
	private int intWorkDay, intMA
	private ResultSet resultSet, resultSet2 = null

	public CheckAkumSama() {


	}

	public String CheckAkumSama2 (String strNama, String strTanggal, String strBroker, int intHariMundur) throws IOException {

		//intHariMundur: pengecekan bbrp hari sebelumnya


		Connection connection = null
		Statement statement = null

		AccessDatabase oAccessDatabase = new AccessDatabase()
		connection = oAccessDatabase.OpenCon()
		statement = connection.createStatement()

		int intTemp1 = 0
		int intLotValue
		String strBuyer1, strBuyer2, strBuyer3, strHasil
		String strTanggalAwal = strTanggal

		strSQL = "select top "+ intHariMundur +" nama, tanggal, Buyer1, BuyerVal1, Buyer2, BuyerVal2, Buyer3, BuyerVal3 from tblBrokSum " +
				"where nama = '"+ strNama +"' and tanggal <= #"+ strTanggal +"# order by tanggal desc "
		System.out.println("strSQL2: " + strSQL)

		resultSet = statement.executeQuery(strSQL)
		intTemp1 = 0
		while (resultSet.next())
		{
			strBuyer1 = resultSet.getString("Buyer1")
			strBuyer2 = resultSet.getString("Buyer2")
			strBuyer3 = resultSet.getString("Buyer3")
			strTanggal = resultSet.getString("tanggal")

			//System.out.println(strTanggal + " : " + strBuyer1 + ", " + strBuyer2 + ", " + strBuyer3)

			if(strBuyer1 == strBroker ) {
				intTemp1++
				intLotValue = resultSet.getInt("BuyerVal1")
				//System.out.println("intTemp1: " + intTemp1 + ", " + "strBroker: " + strBroker)

			}else if(strBuyer2 == strBroker) {

				intTemp1++
				intLotValue = resultSet.getInt("BuyerVal2")

			}else if(strBuyer3 == strBroker) {

				intTemp1++
				intLotValue = resultSet.getInt("BuyerVal3")

			}

			if(intTemp1 == intHariMundur && intLotValue >= 10000) {
				//System.out.println("Info: 5 Hari Akum Sama by di tanggal: " + d.toString())
				strHasil = "ada info: " + strBroker + " (" + intLotValue + ") lot " +  intHariMundur + " Hari Akum " + strNama + " di tanggal: " + strTanggal

				
				// cek kenaikan harga N hari setelah ada akum berturut2
				AccessCekChangePrice oAccessCekChangePrice = new AccessCekChangePrice()
				oAccessCekChangePrice.CekChangePrice2(strNama, strTanggalAwal)
				//System.out.println("Naik ")
				
				//FR: save ke history

			}

		}

		//return strNama + ", " + strTanggal + ", " + strBroker
		return strHasil

	}
}
