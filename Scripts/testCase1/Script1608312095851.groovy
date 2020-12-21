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
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import groovy.json.JsonOutput as JsonOutput
import internal.GlobalVariable as GlobalVariable

ArrayList<Integer> jsonLoaded = new ArrayList<Integer>()

for (def i = 1; i <= findTestData('countriesName').getRowNumbers(); i++) {
    // WebUI.navigateToUrl('https://api.covid19api.com/country/${GlobalVariable.country}')
    //WebUI.navigateToUrl('https://api.covid19api.com/country/' + findTestData('countriesName').getValue(1, i))
    GlobalVariable.Country = findTestData('countriesName').getValue(1, i)

    def result = WS.sendRequest(findTestObject('Object Repository/getJson'))

    def json = JsonOutput.toJson(result)

    new File(('.\\JsonFiles\\' + GlobalVariable.Country) + '.json').write(json)

    WebUI.openBrowser('https://app.flourish.studio/visualisation/4724976/edit?')

    Thread.sleep(1000)

    WebUI.navigateToUrl('https://app.flourish.studio/visualisation/4724976/')

    Thread.sleep(1000)

    WebUI.click(findTestObject('Object Repository/Page_Untitled visualisation  Flourish/button_Duplicate and edit'))

    Thread.sleep(1000)

    WebUI.click(findTestObject('Object Repository/Page_Flourish  Data Visualisation  Storytelling/a_Signin'))

    Thread.sleep(1000)

    WebUI.setText(findTestObject('Object Repository/Page_Flourish  Data Visualisation  Storytelling/input_Email_email'), 
        'sesilva.93@gmail.com')

    WebUI.setEncryptedText(findTestObject('Object Repository/Page_Flourish  Data Visualisation  Storytelling/input_Password_password'), 
        '/9KjW0C3FQ0IRi+Llm2lxg==')

    Thread.sleep(1000)

    WebUI.click(findTestObject('Object Repository/Page_Flourish  Data Visualisation  Storytelling/input_Password_btn'))

    Thread.sleep(1000)

    WebUI.click(findTestObject('Object Repository/Page_Untitled visualisation  Flourish/button_Data'))

    Thread.sleep(1000)

    WebUI.click(findTestObject('Object Repository/Page_Untitled visualisation  Flourish/span_Upload data'))

    Thread.sleep(1000)

	CustomKeywords.'test.fileUploadHelps.uploadFile'(findTestObject(null), 'C:\\Users\\se_si\\Desktop\\Fichas_semanales\\Proyecto_ciclo\\PI_Katalon\\JsonFiles\\'+ GlobalVariable.Country+'.json')
	
	Thread.sleep(1000)

    WebUI.click(findTestObject('Object Repository/Page_Untitled visualisation  Flourish/div_Okay'))

    Thread.sleep(1000)

    WebUI.closeBrowser()

    Thread.sleep(2000) //WebUI.closeBrowser()
}


