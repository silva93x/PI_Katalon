import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject

import com.kms.katalon.core.checkpoint.Checkpoint as Checkpoint
import com.kms.katalon.core.model.FailureHandling as FailureHandling
import com.kms.katalon.core.testcase.TestCase as TestCase
import com.kms.katalon.core.testdata.TestData as TestData
import com.kms.katalon.core.testobject.TestObject as TestObject

import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile

import internal.GlobalVariable as GlobalVariable

import com.kms.katalon.core.annotation.BeforeTestCase
import com.kms.katalon.core.annotation.BeforeTestSuite
import com.kms.katalon.core.annotation.AfterTestCase
import com.kms.katalon.core.annotation.AfterTestSuite
import com.kms.katalon.core.context.TestCaseContext
import com.kms.katalon.core.context.TestSuiteContext as TestSuiteContext

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import java.text.DateFormat as DateFormat
import java.text.SimpleDateFormat as SimpleDateFormat
import java.util.Date as Date
import com.kms.katalon.core.testdata.CSVData as CSVData
import com.kms.katalon.core.testdata.reader.CSVSeparator as CSVSeparator
import sun.misc.BASE64Encoder
import com.jcraft.jsch.*
import com.jcraft.jsch.ChannelSftp as ChannelSftp
import com.kms.katalon.core.util.KeywordUtil

import com.kms.katalon.core.configuration.RunConfiguration

class Actions {
	
	def ClimpiarDescargas() {			   
		//Borra los archivos export.csv de la carpeta Download
		def comando = "powershell.exe del '.\\Data Files\\*.csv'"
	   
		comando.execute()
		Thread.sleep(5000)
		CustomKeywords.'keywords.Utils.Log'("Archivos borrados...")
	   
	}

	/**
	 * Executes before every test suite starts.
	 * @param testSuiteContext: related information of the executed test suite.
	 * Descarga los ficheros que se utilizarán para la ejecución
	 */
	@BeforeTestSuite
	def downloadFileSFTP(TestSuiteContext testSuiteContext){
		// Util para probar en desarrollo
		if (GlobalVariable.downloadSFTP==false) return;
		
		String nameTest = testSuiteContext.getTestSuiteId()
		nameTest = nameTest.substring(nameTest.indexOf("/")+1, nameTest.length())
		java.util.Properties config = new java.util.Properties()
		config.put "StrictHostKeyChecking", "no"
		
		def csvTestMap = [ // Relacion de TestSuite con sus respectivos ficheros
			"Pruebas Unitarias/Test_RegisterEmployee" 	: GlobalVariable.fileRegisterEmployee ]
		
		ClimpiarDescargas()
		
		/*JSch jsch = new JSch();
		Session sess = jsch.getSession GlobalVariable.userSFTP, GlobalVariable.servidorSFTP, 22	
		
		String csvFileName 
		try{
			sess.with {
				CustomKeywords.'keywords.Utils.Log'("Conectando a FTP...")
				setConfig config
				setPassword GlobalVariable.passSFTP
				connect()
				Channel chan = openChannel "sftp"
				chan.connect()
				CustomKeywords.'keywords.Utils.Log'("Conexión FTP realizada...")
				ChannelSftp sftpChannel = (ChannelSftp) chan;
				
				if (nameTest.equals("Test_Complete")){
					CustomKeywords.'keywords.Utils.Log'("Test Suite Test_Complete")
					csvTestMap.each { test , csvName ->
						CustomKeywords.'keywords.Utils.Log'("Descarga fichero $csvName")
						sftpChannel.get( GlobalVariable.rutaFicheros + csvName, "./Data Files" )
					}
				} else if (nameTest.equals("Test_RegisterEPIC")) {
					CustomKeywords.'keywords.Utils.Log'("Test Suite Test_RegisterEPIC")
					csvTestMap.each { test , csvName ->
						if (test.indexOf("Register")>-1){ // Todos los register (EPIC)
							CustomKeywords.'keywords.Utils.Log'("Descarga fichero $csvName")
							sftpChannel.get( GlobalVariable.rutaFicheros + csvName, "./Data Files" )
						}
					}
				} else if (nameTest.equals("Test_SFs")) { // Todos los test de SuccessFactors
					CustomKeywords.'keywords.Utils.Log'("Test Suite Test_SFs")
					csvTestMap.each { test , csvName ->
						if (test.indexOf("Register")<0){ // Todos los que NO son register (EPIC)
							CustomKeywords.'keywords.Utils.Log'("Descarga fichero $csvName")
							sftpChannel.get( GlobalVariable.rutaFicheros + csvName, "./Data Files" )
						}
					}
				} else {
					CustomKeywords.'keywords.Utils.Log'("Test Suite $nameTest")
					csvFileName = csvTestMap[ nameTest ]
					CustomKeywords.'keywords.Utils.Log'("Descarga fichero $csvFileName")
					sftpChannel.get( GlobalVariable.rutaFicheros + csvFileName, "./Data Files" )
				}
				
				sftpChannel.exit()
				chan.disconnect()
				disconnect()
				CustomKeywords.'keywords.Utils.Log'("Desconexion FTP...")
			}
		}catch(Exception e){
			GlobalVariable.errorSFTP = true
			if(e.getMessage().equals("File not found")){
				GlobalVariable.mensajeSFTP = "El fichero " + csvFileName + " no existe en la ruta "+ GlobalVariable.rutaFicheros +" del servidor SFTP. "	
			}else{
				GlobalVariable.mensajeSFTP = "No se ha podido conectar el servidor SFTP " + GlobalVariable.servidorSFTP +" , revisar credenciales"
			}
			
			CustomKeywords.'keywords.Utils.Log'(GlobalVariable.mensajeSFTP)
		}*/		
			
	}
	/**
	 * Executes before every test suite starts.
	 * @param testSuiteContext: related information of the executed test suite.
	 * Se genera el fichero donde se reportarán los resultados
	 */
	@BeforeTestSuite
	def AcrearInforme(TestSuiteContext testSuiteContext) {		
		String nameTest = testSuiteContext.getTestSuiteId()
		nameTest = nameTest.substring(nameTest.indexOf("/")+1, nameTest.length())
		Date fechaActual = new Date()
		DateFormat formatoFecha = new SimpleDateFormat('dd-MM-yyyy-HH.mm.ss', new Locale('es', 'ES'))
		String fecha = formatoFecha.format(fechaActual)
		String nombreInforme

		FileInputStream file = new FileInputStream (new File(".\\Data Files\\Plantilla_InformePruebas_WS.xlsx"))
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheet("Resultados")
		if (nameTest.equals("Test_RegisterEmployee")){
			nombreInforme = "Resultados_EjecucionWSEpic" + fecha + ".xlsx"
		}else{
			nombreInforme = "Resultados_EjecucionWS" + fecha + ".xlsx"
		}		
		
		GlobalVariable.informe = nombreInforme
		
		File reportFile = new File(".\\SpecificReports\\" + GlobalVariable.informe)
		FileOutputStream outFile = new FileOutputStream(reportFile);
		workbook.write(outFile);
		
		GlobalVariable.informeSalida = reportFile.getAbsolutePath()
		CustomKeywords.'keywords.Utils.Log'(GlobalVariable.informeSalida)
		outFile.close();
		Thread.sleep(10000)
	}
	
	@BeforeTestSuite
	def BencodeAutorizacion(TestSuiteContext testSuiteContext){
		String EPICtoken, WStoken
		BASE64Encoder base = new BASE64Encoder()
		EPICtoken = (GlobalVariable.userEpic + ':') + GlobalVariable.passEpic			
		WStoken = (GlobalVariable.userWS + ':') + GlobalVariable.passWS
		GlobalVariable.EPIC_Authorization = 'Basic ' + base.encode(EPICtoken.getBytes())
		GlobalVariable.WS_Authorization = 'Basic ' + base.encode(WStoken.getBytes())
		
		/* 16/09/2020 - Soluciona mensaje 'Request execution failed due to missing or invalid XSRF token' */
		/* 18/09/2020 - Actualizado ! Hay que integrar las Cookies también para las llamadas POST de EPIC */
		def result = CustomKeywords.'keywords.Utils.GetTokenAndCookie'()
		GlobalVariable.tokenEPIC = result["token"]		
		GlobalVariable.cookiesEPIC = [ result["cookie1"], result["cookie2"] ,result["cookie3"] ]
	}
	
	/**
	 * Executes before every test suite starts.
	 * @param testSuiteContext: related information of the executed test suite.
	 * Se guardan los empleados que existen para determinar como informar el campo originalStartDate de EmpEmployment y event-reason de EmpJob
	 */
	/*@BeforeTestSuite
	def CguardarEmpleados(TestSuiteContext testSuiteContext) {
		if(GlobalVariable.errorSFTP==false){
		
			String nameTest = testSuiteContext.getTestSuiteId()
		
			nameTest = nameTest.substring(nameTest.indexOf("/")+1, nameTest.length())
		
			//Se ejecuta el proceso según el TestSuite seleccionado
			if(nameTest.equals("Test_Employment") || nameTest.equals("Test_EmpJob") || nameTest.equals("Test_Complete")){
				List<String> rutaFichero = []
				List<String> fileEmpEmployment = []
				List<String> fileEmpJob = []
				
				//Se añade los ficheros a tratar según el TestSuite seleccionado
				if(nameTest.equals("Test_Employment")){
					rutaFichero.add('.\\Data Files\\' + GlobalVariable.fileEmployment)
				}else{
					if(nameTest.equals("Test_EmpJob")){
						rutaFichero.add('.\\Data Files\\' + GlobalVariable.fileEmpJob)
					}else{
						rutaFichero.add('.\\Data Files\\' + GlobalVariable.fileEmployment)
						rutaFichero.add('.\\Data Files\\' + GlobalVariable.fileEmpJob)
					}
				}
			
				CSVData csvData
				int filas
				String id
			
				//Se comprueban la existencia de los empleados de los ficheros EmpEmployment y EmpJob
				for(int i=0;i < rutaFichero.size();i++){
					
					csvData = new CSVData(rutaFichero.get(i), true, CSVSeparator.COMMA)
					
					if (GlobalVariable.contadorTotal == -1) {
						filas = csvData.getRowNumbers()
					} else {
						filas = GlobalVariable.contadorTotal
					}
					
					//filas = csvData.getRowNumbers()
					
					//Se comprueba la existencia de to
					for(int x=1; x<filas; x++){
					
						id = csvData.getValue("user-id", x + 1)
						
						//Llamada al WS para verificar sí existe el usuario
						String url = GlobalVariable.urlApi + "User?\$select=userId&\$filter=userId eq '"+ id + "'"
						
						HttpURLConnection connection = (HttpURLConnection) new URL(url).openConnection()
						
						//Se informa el tipo de petición y la cabecera
						connection.setRequestMethod('GET')
						connection.setRequestProperty('Content-Type', 'application/tom-xml')
						connection.setRequestProperty('Authorization', GlobalVariable.Authorization)
						
						// Se verifica el código de respuesta y se obtiene el cuerpo de la respuesta
						int responseCode = connection.getResponseCode()
						InputStream inputStream
						if (200 <= responseCode && responseCode <= 299) {
							inputStream = connection.getInputStream()
						} else {
							inputStream = connection.getErrorStream()
						}
			
						//Se trata la respuesta para finalmente tener una cadena
						BufferedReader a = new BufferedReader(
								new InputStreamReader(
								inputStream));
				
						StringBuilder response = new StringBuilder()
						String currentLine;
				
						while ((currentLine = a.readLine()) != null)
							response.append(currentLine)
				
						a.close()
				
						String b = response.toString()
				
						String busqueda = "<d:userId>"+id+"</d:userId>"
						
						//Sí existe la etiqueta con el id del empleado significa que el usuario existe
						if(b.indexOf(busqueda)!=-1) {
							if(rutaFichero.get(i).equals('.\\Data Files\\' + GlobalVariable.fileEmployment)){
								fileEmpEmployment.add(id)
								GlobalVariable.existentesEmployment = fileEmpEmployment 
							}else{
								fileEmpJob.add(id)
								GlobalVariable.existentesJob = fileEmpJob
								
							}
						}
						CustomKeywords.'keywords.Utils.Log'("Procesado el registro: " + x)
					}
				}//End del for para cada uno de los ficheros
			}//End if según el nombre del TestSuite
			CustomKeywords.'keywords.Utils.Log'("Fin de la carga de empleados existentes")
		}//End if de la va
	}*/
}