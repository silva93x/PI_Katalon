//import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject

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


//import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle

//import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.VerticalAlignment
//import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.xssf.usermodel.XSSFCellStyle as XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.junit.Before

import java.util.List

import java.text.DateFormat as DateFormat
import java.text.SimpleDateFormat as SimpleDateFormat
import java.util.Date as Date
import org.apache.poi.hssf.usermodel.HSSFCell
import org.apache.poi.hssf.usermodel.HSSFCellStyle
import org.apache.poi.hssf.usermodel.HSSFSheet
import com.kms.katalon.core.configuration.RunConfiguration
import internal.GlobalVariable


public class Excel {

	@Keyword
	def static void writeExcel(List<String> MisDatos){

		Date fechaActual = new Date()
		DateFormat formatoFecha = new SimpleDateFormat('dd/MM/yyyy - HH:mm:ss', new Locale('es', 'ES'))
		String Fecha = formatoFecha.format(fechaActual)

		int iCell = 1
		int iRow
		String RutaProyecto = RunConfiguration.getProjectDir()
		RutaProyecto = RutaProyecto.replace("/", "\\\\")

		FileInputStream file = new FileInputStream (new File(RutaProyecto + "\\SpecificReports\\" + GlobalVariable.informe))
		XSSFWorkbook workbook = new XSSFWorkbook(file)
		XSSFSheet sheet = workbook.getSheet("Resultados")
		XSSFCellStyle cellStyle
		//Linea para actualizar las fórmulas del excel
		workbook.setForceFormulaRecalculation(true)

		//--------------Write data to excel--------------
		XSSFRow oRow
		XSSFCell oCell


		//--------------Informar a Entorno--------------
		oRow = sheet.getRow(3)
		oCell = oRow.getCell(3)
		oCell.setCellValue(GlobalVariable.urlWS)
		oRow = sheet.getRow(iRow)
		oRow.createCell(iCell)

		//--------------Informar a Fecha----------------
		oRow = sheet.getRow(4)
		oCell = oRow.getCell(3)
		oCell.setCellValue(Fecha)
		oRow = sheet.getRow(iRow)
		oRow.createCell(iCell)

		//Informar a Documetación Auxiliar

		//--------------Informar Nombre WS-------------

		oRow = sheet.getRow(GlobalVariable.indice)

		if (oRow == null){
			sheet.createRow(iRow)
			oRow = sheet.getRow(iRow)
			sheet.createRow(iCell)

		}
		sheet.createRow(iRow)
		sheet.createRow(iCell)
		oCell = oRow.getCell(iCell)
		oRow.createCell(iCell)
		oCell = oRow.getCell(iCell)
		sheet.createRow(iRow).createCell(iCell)
		//Le asignamos estilos a nuestro excel, le ponemos todos los bordes
		cellStyle = workbook.createCellStyle()
		cellStyle.setBorderBottom(BorderStyle.THIN)
		cellStyle.setBorderLeft(BorderStyle.THICK)
		cellStyle.setBorderRight(BorderStyle.THIN)
		cellStyle.setBorderTop(BorderStyle.THIN)
		oCell.setCellStyle(cellStyle)
		oCell.setCellValue(MisDatos.get(0));

		//--------------Informar UserID--------------
		//if(oRow == null){
		//	sheet.createRow(iRow);
		//	oRow = sheet.getRow(iRow);
		//}
		oRow = sheet.getRow(GlobalVariable.indice)
		oCell = oRow.getCell(iCell + 1);
		if(oCell == null ){
			oRow.createCell(iCell + 1);
			oCell = oRow.getCell(2);
			//Le asignamos estilos a nuestro excel, le ponemos todos los bordes
			cellStyle = workbook.createCellStyle();
			cellStyle.setBorderBottom(BorderStyle.THIN)
			cellStyle.setBorderLeft(BorderStyle.THIN)
			cellStyle.setBorderRight(BorderStyle.THIN)
			cellStyle.setBorderTop(BorderStyle.THIN)
			//cellStyle.setAlignment(CellStyle.ALIGN_CENTER)
			cellStyle.setAlignment(HorizontalAlignment.CENTER)
			oCell.setCellStyle(cellStyle)
		}
		oCell.setCellValue(MisDatos.get(1));
		//--------------Informar Código Respuesta--------------
		//		if(oRow == null){
		//			sheet.createRow(iRow);
		//			oRow = sheet.getRow(iRow)
		//		}
		oRow = sheet.getRow(GlobalVariable.indice)
		oCell = oRow.getCell(iCell + 2)
		if(oCell == null ){
			oRow.createCell(iCell + 2)
			oCell = oRow.getCell(3)
			//Le asignamos estilos a nuestro excel, le ponemos todos los bordes
			cellStyle = workbook.createCellStyle()
			cellStyle.setBorderBottom(BorderStyle.THIN)
			cellStyle.setBorderLeft(BorderStyle.THIN)
			cellStyle.setBorderRight(BorderStyle.THIN)
			cellStyle.setBorderTop(BorderStyle.THIN)
			cellStyle.setAlignment(HorizontalAlignment.CENTER)
			oCell.setCellStyle(cellStyle)
		}
		oCell.setCellValue(MisDatos.get(2))
		//--------------Informar Resultado--------------
		//		if(oRow == null){
		//			sheet.createRow(iRow)
		//			oRow = sheet.getRow(iRow)
		//		}
		oRow = sheet.getRow(GlobalVariable.indice)
		oCell = oRow.getCell(iCell + 3)
		if(oCell == null ){
			oRow.createCell(iCell + 3)
			oCell = oRow.getCell(4)
			//Le asignamos estilos a nuestro excel, le ponemos todos los bordes y alineamos al centro
			cellStyle = workbook.createCellStyle()
			cellStyle.setBorderBottom(BorderStyle.THIN)
			cellStyle.setBorderLeft(BorderStyle.THIN)
			cellStyle.setBorderRight(BorderStyle.THIN)
			cellStyle.setBorderTop(BorderStyle.THIN)
			cellStyle.setAlignment(HorizontalAlignment.CENTER)
			oCell.setCellStyle(cellStyle)
		}
		oCell.setCellValue(MisDatos.get(3))
		//--------------Informar Respuesta WS Código--------------
		//		if(oRow == null){
		//			sheet.createRow(iRow)
		//			oRow = sheet.getRow(iRow)
		//		}
		oRow = sheet.getRow(GlobalVariable.indice)
		oCell = oRow.getCell(iCell + 4)
		if(oCell == null ){
			oRow.createCell(iCell + 4)
			oCell = oRow.getCell(5)
			//Le asignamos estilos a nuestro excel, le ponemos todos los bordes y alineamos al centro
			cellStyle = workbook.createCellStyle()
			cellStyle.setBorderBottom(BorderStyle.THIN)
			cellStyle.setBorderLeft(BorderStyle.THIN)
			cellStyle.setBorderRight(BorderStyle.THIN)
			cellStyle.setBorderTop(BorderStyle.THIN)
			cellStyle.setAlignment(HorizontalAlignment.CENTER)
			oCell.setCellStyle(cellStyle)
		}
		oCell.setCellValue(MisDatos.get(4))
		//--------------Informar Respuesta WS Estado--------------
		//		if(oRow == null){
		//			sheet.createRow(iRow)
		//			oRow = sheet.getRow(iRow)
		//		}
		oRow = sheet.getRow(GlobalVariable.indice)
		oCell = oRow.getCell(iCell + 5)
		if(oCell == null ){
			oRow.createCell(iCell + 5)
			oCell = oRow.getCell(6)
			//Le asignamos estilos a nuestro excel, le ponemos todos los bordes y alineamos al centro
			cellStyle = workbook.createCellStyle()
			cellStyle.setBorderBottom(BorderStyle.THIN)
			cellStyle.setBorderLeft(BorderStyle.THIN)
			cellStyle.setBorderRight(BorderStyle.THIN)
			cellStyle.setBorderTop(BorderStyle.THIN)
			cellStyle.setAlignment(HorizontalAlignment.CENTER)
			oCell.setCellStyle(cellStyle)
		}
		oCell.setCellValue(MisDatos.get(5))
		//--------------Informar Respuesta WS Exception--------------
		//		if(oRow == null){
		//			sheet.createRow(iRow)
		//			oRow = sheet.getRow(iRow)
		//		}
		oRow = sheet.getRow(GlobalVariable.indice)
		oCell = oRow.getCell(iCell + 6)
		//if(oCell == null ){
		oRow.createCell(iCell + 6)
		oCell = oRow.getCell(7)
		//Le asignamos estilos a nuestro excel, le ponemos todos los bordes
		cellStyle = workbook.createCellStyle()
		cellStyle.setBorderBottom(BorderStyle.THIN)
		cellStyle.setBorderLeft(BorderStyle.THIN)
		cellStyle.setBorderRight(BorderStyle.THIN)
		cellStyle.setBorderTop(BorderStyle.THIN)
		cellStyle.setAlignment(HorizontalAlignment.CENTER)
		cellStyle.setVerticalAlignment(VerticalAlignment.TOP)
		cellStyle.setWrapText(true)
		oCell.setCellStyle(cellStyle)
		//}
		oCell.setCellValue(MisDatos.get(6))
		//----------------Informar Observaciones-------------------
		//		if(oRow == null){
		//			sheet.createRow(iRow)
		//			oRow = sheet.getRow(iRow)
		//		}
		oRow = sheet.getRow(GlobalVariable.indice)
		oCell = oRow.getCell(iCell + 7)
		//if(oCell == null ){
		oRow.createCell(iCell + 7)
		oCell = oRow.getCell(8)
		//Combinamos celdas, pero esta obsoleto.
		//sheet.addMergedRegion (new CellRangeAddress ( 8 , 8 , 8 , 11 ));
		//Le asignamos estilos a nuestro excel, le ponemos todos los bordes y alineado a la izquierda
		cellStyle = workbook.createCellStyle()
		cellStyle.setBorderBottom(BorderStyle.THIN)
		cellStyle.setBorderLeft(BorderStyle.THIN)
		cellStyle.setBorderRight(BorderStyle.THICK)
		cellStyle.setBorderTop(BorderStyle.THIN)
		cellStyle.setAlignment(HorizontalAlignment.LEFT)
		cellStyle.setVerticalAlignment(VerticalAlignment.TOP)
		cellStyle.setWrapText(true)
		oCell.setCellStyle(cellStyle)
		//}
		oCell.setCellValue(MisDatos.get(7))


		//--------------Guardar excel--------------
		String nombreInforme =  GlobalVariable.informe
		FileOutputStream outFile = new FileOutputStream( new File(".\\SpecificReports\\" + GlobalVariable.informe))
		workbook.write(outFile)
		//outFile.flush()
		outFile.close()
		GlobalVariable.indice = GlobalVariable.indice + 1
	}

}
