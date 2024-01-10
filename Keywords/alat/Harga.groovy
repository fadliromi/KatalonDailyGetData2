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
import com.kms.katalon.core.testobject.SelectorMethod
import org.openqa.selenium.Keys as Keys
import internal.GlobalVariable

public class Harga {


	private String[] strTemp1 = new String[3]


	public Harga(){
	}

	public String[] AmbilData (String strKode, Integer intTick) throws IOException {


		//buka tab baru
		WebUI.executeJavaScript('window.open();', [])
		WebUI.switchToWindowIndex(1)
		WebUI.navigateToUrl("https://stockbit.com/symbol/" + strKode)
		WebUI.delay(2)
		WebUI.executeJavaScript("document.body.style.zoom='20%'", null)
		WebUI.maximizeWindow()

		WebUI.scrollToElement(findTestObject('Scalping/label_General_AvgPrice'), 0)
		strTemp1[0] = WebUI.getText(findTestObject('Scalping/label_General_AvgPrice')) //AVG price


		// Tentukan fraksi harga
		Integer intHarga = strTemp1[0].toInteger()
		if(intHarga < 200) {
			System.out.println("Fraksi harga Rp.1")
			//intHarga = intHarga - (4 * intHarga)
			strTemp1[1] = intHarga - (intTick * 1)
			System.out.println("AVG: " + intHarga + ", -" + intTick + "tick: " + strTemp1[1] )

		}else if(intHarga >= 200 && intHarga < 500) {
			System.out.println("Fraksi harga Rp.2")
			strTemp1[1] = intHarga - (intTick * 2)
			System.out.println("AVG: " + intHarga + ", -" + intTick + "tick: " + strTemp1[1] )

		}else if(intHarga >= 500 && intHarga < 2000) {
			System.out.println("Fraksi harga Rp.5")
			strTemp1[1] = intHarga - (intTick * 5)
			System.out.println("AVG: " + intHarga + ", -" + intTick + "tick: " + strTemp1[1] )

		}else {
			System.out.println("Fraksi harga tidak ditemukan")
			strTemp1[1] = null
		}
		// Tentukan fraksi harga

		WebUI.setViewPortSize(800, 650)
		WebUI.closeWindowIndex(1)
		WebUI.switchToWindowIndex(0)
		return strTemp1
	}
}


public class AlertPrice {

	//private String strXpath

	public AlertPrice(){

	}

	public AddAlert (String strKode, String strHarga) throws IOException {
		
		TestObject to = findTestObject('Scalping/button_SubScreener_Favorites')
		String strXpath

		WebUI.click(findTestObject('Scalping/button_SubSideMenu_AlertPrice'))


		//cek apakah alert sudah ada
		strXpath = "//*[@class = 'ant-row' and (contains(text(), '" + strKode + "') or contains(., '" + strKode + "'))]"
		to.setSelectorValue(SelectorMethod.XPATH, strXpath)
		to.setSelectorMethod(SelectorMethod.XPATH)

		//WebUI.verifyElementNotVisible(findTestObject, FailureHandling.STOP_ON_FAILURE)

		if(WebUI.verifyElementVisible(to, FailureHandling.OPTIONAL) == false) {
			//jika belum ada alert
			System.out.println("Ada info: Tambah alert baru " + strKode)

			WebUI.click(findTestObject('Scalping/label_SubAlertPrice_AddAlert'))
			WebUI.delay(2)
			WebUI.setText(findTestObject('Scalping/label_SubAlertPrice_AddAlert_StockName'), strKode)
			WebUI.delay(2)

			strXpath = "//*[@type = 'grey' and (text() = '" + strKode + "' or . = '" + strKode + "')]"
			to.setSelectorValue(SelectorMethod.XPATH, strXpath)
			to.setSelectorMethod(SelectorMethod.XPATH)
			WebUI.click(to)

			WebUI.click(findTestObject('Scalping/button_SubAlertPrice_AddAlert_Condition'))
			WebUI.click(findTestObject('Scalping/button_SubAlertPrice_AddAlert_Condition2'))
			WebUI.setText(findTestObject('Scalping/input_SubAlertPrice_AddAlert_Value'), strHarga)
			WebUI.click(findTestObject('Scalping/button_SubAlertPrice_AddAlert_Create'))

		}else {
			System.out.println("Ada info: Edit existing alert "  + strKode)

			//edit existing alert
			strXpath = "//*[@class = 'ant-row' and (contains(text(), '" + strKode + "') or contains(., '" + strKode + "'))]"
			to.setSelectorValue(SelectorMethod.XPATH, strXpath)
			to.setSelectorMethod(SelectorMethod.XPATH)
			WebUI.click(to)

			WebUI.sendKeys(findTestObject('Scalping/input_SubAlertPrice_AddAlert_Value'), Keys.chord(Keys.CONTROL, 'a'))
			WebUI.sendKeys(findTestObject('Scalping/input_SubAlertPrice_AddAlert_Value'), Keys.chord(Keys.BACK_SPACE))

			WebUI.setText(findTestObject('Scalping/input_SubAlertPrice_AddAlert_Value'), strHarga)
			WebUI.click(findTestObject('Scalping/button_SubAlertPrice_AddAlert_Create'))

		}


		WebUI.click(findTestObject('Scalping/button_SubSideMenu_AlertPrice'))
	}
}




