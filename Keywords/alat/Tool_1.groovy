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


import internal.GlobalVariable

public class Tool_1 {


	public Tool_1(String BndrValue, String BndrAVG){

	}
	
}

public class CekChangePrice {
	
			
	
		public CekChangePrice(){
			
		}
		
		public CekChangePrice2 (String strNama, String strTgl) throws IOException {
			
			// cek persentase kenaikan N hari ke depan
			String strSQL, strTanggal
			Double dblChange
			strTanggal = strTgl
			
			Connection connection = null
			Statement statement = null
			ResultSet resultSet
	
			AccessDatabase oAccessDatabase = new AccessDatabase()
			connection = oAccessDatabase.OpenCon()
			statement = connection.createStatement()
			
			//strSQL = "select top3, nama, tanggal, change from TblBrokSum where nama = '" + strNama + "' and tanggal >= #2023-12-04#"
			strSQL = "select top3, nama, tanggal, change from TblBrokSum where nama = '" + strNama + "' and tanggal >= #"+ strTgl +"#"
			System.out.println("strSQL2: " + strSQL)
			
			resultSet = statement.executeQuery(strSQL)
			
			while (resultSet.next())
			{
				dblChange = resultSet.getDouble("change")
				strTanggal = resultSet.getDate("tanggal")
				
				
				
				if(dblChange >= 10) {
					System.out.println("Info Naik Setelah Akum Sama: Tanggal: " + strTanggal + ", " + strNama + " Naik " + dblChange + "%")
					break
				}
				
			}
			
			
		}
		
	}