package org.svetikov.reportlaboratorymdf


import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.springframework.boot.CommandLineRunner
import org.springframework.boot.autoconfigure.SpringBootApplication
import org.springframework.boot.runApplication
import java.io.IOException
import kotlin.jvm.Throws

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.time.LocalDateTime
import kotlin.math.pow
import kotlin.math.roundToInt
import kotlin.math.sqrt
import kotlin.reflect.full.findAnnotation
import kotlin.reflect.full.functions


@SpringBootApplication
class ReportLaboratoryMdfApplication : CommandLineRunner {

    val a: Int = 5

    @Throws(IOException::class, NullPointerException::class)
    override fun run(vararg args: String?) {

//todo init list from imal
        val listDicke = mutableListOf<String>()
        val listRohdichte = mutableListOf<String>()
        val listQuerzug = mutableListOf<String>()
        val listAbhebe = mutableListOf<String>()
//todo calling read data from imal
        readTableFromIMAL(
            "Dicke" to listDicke,
            "Rohdichte" to listRohdichte,
            "Querzug" to listQuerzug,
            "Abhebe" to listAbhebe
        )
//todo create table customer
        makeTableCustomer(listDicke, listRohdichte, listQuerzug, listAbhebe)

        //todo-----------------------------ReflektionAPI---------------------------Anatattion
//        val classs = ReportLaboratoryMdfApplication::class
//
//        val anat = classs.findAnnotation<ListName>()
//        println(anat?.name+ "   hfdbhdf")
//
//
//        listDicke.map { it + " dicke" }.forEach { println(it) }
//        listRohdichte.map { it + "  roh" }.forEach { println(it) }
//        listQuerzug.map { it + " quez" }.forEach { println(it) }
//        listAbhebe.map { it + "  abhebe" }.forEach { println(it) }

    }


}

fun makeTableCustomer(vararg listDataValueFromIMAL: MutableList<String>) {


    Customer(1, "fg", "gf", 5).also { println(it) }
    val workBook = XSSFWorkbook()
    val createHelper = workBook.creationHelper

    val sheet = workBook.createSheet("Customer")

    sheet.setColumnWidth(0, 1000)
    sheet.setColumnWidth(1, 2000)
    sheet.setColumnWidth(2, 4000)
    sheet.setColumnWidth(3, 4000)
    sheet.setColumnWidth(4, 2000)


    //todo header title
    val headerTitle = workBook.createCellStyle()
    headerTitle.setVerticalAlignment(VerticalAlignment.CENTER)
    headerTitle.setAlignment(HorizontalAlignment.CENTER)
    headerTitle.setFillForegroundColor(IndexedColors.AQUA.index)
    headerTitle.setFillPattern(FillPatternType.SOLID_FOREGROUND)
    val rowTitle = sheet.createRow(1)
    val cellTitle = rowTitle.createCell(1)
    cellTitle.setCellValue("Fiberboard")
    sheet.addMergedRegion(CellRangeAddress(0, 1, 1, 10))
    cellTitle.cellStyle = headerTitle
//todo font
    val fontForPLCTable = workBook.createFont()
    fontForPLCTable.setFontHeight(9.0)
    val fontTable = workBook.createFont()
    fontTable.setFontHeight(10.0)
    //todo general style
    val generalStyle = workBook.createCellStyle()
    generalStyle.setVerticalAlignment(VerticalAlignment.TOP)
    generalStyle.setAlignment(HorizontalAlignment.CENTER)
    generalStyle.setBorderTop(BorderStyle.THIN)
    generalStyle.setBorderBottom(BorderStyle.THIN)
    generalStyle.setBorderLeft(BorderStyle.THIN)
    generalStyle.setBorderRight(BorderStyle.THIN)
    generalStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.index)
    generalStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND)


    val generalStyleTable = workBook.createCellStyle()
    generalStyleTable.setVerticalAlignment(VerticalAlignment.TOP)
    generalStyleTable.setAlignment(HorizontalAlignment.CENTER)
    generalStyleTable.setBorderTop(BorderStyle.THIN)
    generalStyleTable.setBorderBottom(BorderStyle.THIN)
    generalStyleTable.setBorderLeft(BorderStyle.THIN)
    generalStyleTable.setBorderRight(BorderStyle.THIN)
    generalStyleTable.setFillForegroundColor(IndexedColors.AQUA.index)
    generalStyleTable.setFillPattern(FillPatternType.SOLID_FOREGROUND)
    generalStyleTable.setFont(fontTable)

    val generalStyleFromPLC_Left = workBook.createCellStyle()
    generalStyleFromPLC_Left.setVerticalAlignment(VerticalAlignment.TOP)
    generalStyleFromPLC_Left.setAlignment(HorizontalAlignment.CENTER)
    generalStyleFromPLC_Left.setBorderLeft(BorderStyle.THIN)
    generalStyleFromPLC_Left.setFont(fontForPLCTable)
    generalStyleFromPLC_Left.setFillForegroundColor(IndexedColors.AQUA.index)
    generalStyleFromPLC_Left.setFillPattern(FillPatternType.SOLID_FOREGROUND)


    val generalStyleFromPLC_Right = workBook.createCellStyle()
    generalStyleFromPLC_Right.setVerticalAlignment(VerticalAlignment.TOP)
    generalStyleFromPLC_Right.setAlignment(HorizontalAlignment.CENTER)
    generalStyleFromPLC_Right.setBorderRight(BorderStyle.THIN)
    generalStyleFromPLC_Right.setFont(fontForPLCTable)
    generalStyleFromPLC_Right.setFillForegroundColor(IndexedColors.SEA_GREEN.index)
    generalStyleFromPLC_Right.setFillPattern(FillPatternType.SOLID_FOREGROUND)


    //todo fusion cell -----------------------------------------------------------------
    sheet.addMergedRegion(CellRangeAddress(2, 3, 1, 2))
// todo init header one
    val listHeaderTextOne = listOf<String>("Charge: \n", "Typ: \n", "Prod.zeit: \n", "Pruf.Nr.: \n")
    val listHeaderTextValuesOne = listOf<String>("123456", "HDF", LocalDateTime.now().toString(), "79563466740")
    //todo text and value 1------------------------------------------------------------------------
    val rowHeaderOne = sheet.createRow(2)

    createHeaderOne(sheet, rowHeaderOne, 0, 1, 1, generalStyle, listHeaderTextOne, listHeaderTextValuesOne)
    //todo text and value 2------------------------------------------------------------------------
    createHeaderOne(sheet, rowHeaderOne, 1, 3, 4, generalStyle, listHeaderTextOne, listHeaderTextValuesOne)
    // todo text and value 3------------------------------------------------------------------------
    createHeaderOne(sheet, rowHeaderOne, 2, 5, 7, generalStyle, listHeaderTextOne, listHeaderTextValuesOne)
    //  todo text and value 4------------------------------------------------------------------------
    createHeaderOne(sheet, rowHeaderOne, 3, 8, 10, generalStyle, listHeaderTextOne, listHeaderTextValuesOne)

    rowHeaderOne.createCell(12).setCellValue("Svetikov")

//todo -----------------------------------------------------------------------------------------------------------------------------------------------------------------------
//todo --Nr--|--Dicke--|--Rohdichte--|--Querzugfestigkeit--|--Abhebe-festigkell--|--Beige-festigkell--|--E-Modul-quer--|--Quellung-24h--|--Rest-feuchte--|Wasser-aufnahme----|--------------------------------------
//todo -----------------------------------------------------------------------------------------------------------------------------------------------------------------------
    val listHeaderTwoText = listOf<String>(
        " Nr ",
        "Dicke",
        "Rohdichte",
        "Querzug-\n festigkeit",
        "Abhebe-\n festigkell",
        "Beige-\n festigkell",
        "E-Modul-\n quer",
        "Quellung- \n 24h",
        "Rest- \n feuchte",
        "Wasser- \n aufnahme"
    )

    val listHeaderTwoSI = listOf<String>(
        "_____", "mm", "kg/m2", "N/mm2", "N/mm2", "N/mm2", "N/mm2", "%", "%", "%"
    )
    val listNumberTwo = listOf<String>("1", "2", "3", "4", "5", "6", "7", "8", "MW", "SOLL", "Max", "Min")


    //todo header text
    val rowHeaderTwoText = sheet.createRow(5)
    var cellHeaderTwoText = rowHeaderTwoText.createCell(1)
    //todo header si
    val rowHeaderTwoSI = sheet.createRow(7)
    var cellHeaderTwoSI = rowHeaderTwoSI.createCell(1)
    //todo table for value from laboratory
    var rowValueFromLaboratory = sheet.createRow(18)
    var cellValueFromLaboratory = rowValueFromLaboratory.createCell(1)

    for (index in listNumberTwo.indices) { //listHeaderTwoText.indices
        //todo create header text

        if (index < 10) {
            cellHeaderTwoText = rowHeaderTwoText.createCell(index + 1)
            cellHeaderTwoText.setCellValue(listHeaderTwoText[index])
            cellHeaderTwoText.cellStyle = generalStyle
        }
        //todo create header SI

        if (index < 10) {
            cellHeaderTwoSI = rowHeaderTwoSI.createCell(index + 1)
            cellHeaderTwoSI.setCellValue(listHeaderTwoSI[index])
            cellHeaderTwoSI.cellStyle = generalStyle

        }

        sheet.autoSizeColumn(index + 1)
        if (index < 10) {
            sheet.addMergedRegion(
                CellRangeAddress(
                    5,
                    6,
                    cellHeaderTwoText.columnIndex,
                    cellHeaderTwoText.columnIndex
                )
            )
        }
        //todo create header value laboratory
        rowValueFromLaboratory = sheet.createRow(index + 8)
        for (i in listHeaderTwoSI.indices) {
//todo create and fulling cells data from imal laboratory
            cellValueFromLaboratory = rowValueFromLaboratory.createCell(i + 1)
            when (i) {
                0 -> cellValueFromLaboratory.setCellValue(listNumberTwo[index])
                1 -> cellValueFromLaboratory.setCellValue(listDataValueFromIMAL[0][index])
                2 -> cellValueFromLaboratory.setCellValue(listDataValueFromIMAL[1][index])
                3 -> cellValueFromLaboratory.setCellValue(listDataValueFromIMAL[2][index])
                4 -> cellValueFromLaboratory.setCellValue(listDataValueFromIMAL[3][index])
                5 -> cellValueFromLaboratory.setCellValue("null")
                6 -> cellValueFromLaboratory.setCellValue("null")
                7 -> cellValueFromLaboratory.setCellValue("null")
                8 -> cellValueFromLaboratory.setCellValue("null")
                9 -> cellValueFromLaboratory.setCellValue("null")
                else -> throw NoSuchElementException("No list with iteration $i")
            }
            cellValueFromLaboratory.cellStyle = generalStyleTable
        }
    }


//todo middle general data (allgemeine)
    val listHeaderData = listOf<String>(
        "Allgemeine Daten",
        "Holzeinsatz",
        "Faserfeuchten", "Beleimung", "Refiner", "Form-/ Pressenstrasse"
    )

    var step = 21 //todo start row
    var bias = 1
    var bias2 = 3
    var rowHeaderDataPLC = sheet.createRow(step)
    for (index in listHeaderData.indices) {
        if (index == 3) { // todo jump to next row
            step += 15
            bias = 1
            bias2 = 3
            rowHeaderDataPLC = sheet.createRow(step)
        }
        val cellHeaderDataPLC = rowHeaderDataPLC.createCell(bias)
        cellHeaderDataPLC.setCellValue(listHeaderData[index])

        if (index == 0 || index == 3) { //todo rice 4 cell first note
            bias2 += 1
            //println("1 $bias // $bias2 //$index")
            sheet.addMergedRegion(CellRangeAddress(step, step, bias, bias2))
            bias = bias2 + 1
            bias2 += 3
            // println("2 $bias // $bias2 //$index")
        }

        if (index != 0 && index != 3) { //todo other cell note
            //   println("3 $bias // $bias2 //$index")
            sheet.addMergedRegion(CellRangeAddress(step, step, bias, bias2))
            bias = bias2 + 1
            bias2 += 3
            //  println("4 $bias // $bias2 //$index")
        }
        cellHeaderDataPLC.cellStyle = generalStyle
    }
    //todo value from PLC

    val listTextValueFromPLC =
        listOf("", "text", "text", "text", "text", "text", "text", "text", "text", "", "", "", "", "")
    val listTextValueFromPLC2 =
        listOf(
            "text",
            "text",
            "text",
            "text",
            "text",
            "text",
            "text",
            "text",
            "text",
            "text",
            "text",
            "text",
            "text",
            "text"
        )
    val listValueFromPLC = listOf(
        "",
        "val",
        "val",
        "val",
        "val" + "mm",
        "val" + "mm",
        "val" + "mm",
        "val" + "mm",
        "val" + "kg/m2",
        "",
        "",
        "",
        "",
        ""
    )
    val listValueFromPLC2 = listOf(
        "val",
        "val",
        "val",
        "val",
        "val" + " mm",
        "0.97" + " mm",
        "val" + " mm",
        "val" + " mm",
        "0.99" + " kg/m2",
        "val",
        "val",
        "val",
        "val",
        "val"
    )
    var biasPLC = 22
    for (index in listValueFromPLC.indices) {

        var rowValueFromPLC = sheet.createRow(biasPLC)
        cellFull(
            sheet, biasPLC, rowValueFromPLC, index, 1, 2, 3, 4,
            generalStyleFromPLC_Left, generalStyleFromPLC_Right, listTextValueFromPLC, listValueFromPLC
        )
        cellFull(
            sheet, biasPLC, rowValueFromPLC, index, 5, 6, 7, 7,
            generalStyleFromPLC_Left, generalStyleFromPLC_Right, listTextValueFromPLC, listValueFromPLC
        )
        cellFull(
            sheet, biasPLC, rowValueFromPLC, index, 8, 9, 10, 10,
            generalStyleFromPLC_Left, generalStyleFromPLC_Right, listTextValueFromPLC2, listValueFromPLC2
        )

        biasPLC += 1

    }
    biasPLC = 37
    for (index in listValueFromPLC.indices) {

        var rowValueFromPLC = sheet.createRow(biasPLC)
        cellFull(
            sheet, biasPLC, rowValueFromPLC, index, 1, 2, 3, 4,
            generalStyleFromPLC_Left, generalStyleFromPLC_Right, listTextValueFromPLC, listValueFromPLC
        )
        cellFull(
            sheet, biasPLC, rowValueFromPLC, index, 5, 6, 7, 7,
            generalStyleFromPLC_Left, generalStyleFromPLC_Right, listTextValueFromPLC2, listValueFromPLC2
        )
        cellFull(
            sheet, biasPLC, rowValueFromPLC, index, 8, 9, 10, 10,
            generalStyleFromPLC_Left, generalStyleFromPLC_Right, listTextValueFromPLC, listValueFromPLC
        )

        biasPLC += 1

    }
    var biasBottom = 51
    for (index in 1..5) {
        val rowBottom = sheet.createRow(biasBottom)
        val cellBottom = rowBottom.createCell(1)
        biasBottom += 1
        cellBottom.cellStyle = generalStyle
    }



    sheet.addMergedRegion(CellRangeAddress(51, biasBottom, 1, 10))


    val fileOut = FileOutputStream("C:/customer.xlsx")
    workBook.write(fileOut)
    fileOut.close()
    workBook.close()
}

fun fullingListInside(vararg pairList: Pair<String, MutableList<String>>) {
//  pairList.forEach { when(it.first){
//      "Dicke"-> it.second.addAll(list1)
//      "Rohdichte"-> it.second.addAll(list2)
//      "Querzug"-> it.second.addAll(list3)
//      else -> throw NoSuchElementException(it.first)
//  } }
}


fun readTableFromIMAL(vararg pairList: Pair<String, MutableList<String>>) {

    val excelFile = FileInputStream(File("C:\\Imal.xlsx"))
    val workBook = XSSFWorkbook(excelFile)
    val sheet = workBook.getSheet("ImalTest")



    pairList.forEach {
        when (it.first) {
            "Dicke" -> it.second.addAll(addAverageMaxMin(getListInTheRange(sheet, "D", "D19", "D29")))
            "Rohdichte" -> it.second.addAll(addAverageMaxMin(getListInTheRange(sheet, "F", "F19", "F28")))
            "Querzug" -> it.second.addAll(addAverageMaxMin(getListInTheRange(sheet, "I", "I19", "I28")))
            "Abhebe" -> it.second.addAll(addAverageMaxMinSkip(2, getListInTheRange(sheet, "I", "I38", "I45")))
            else -> throw NoSuchElementException(it.first)
        }
    }


}

fun getListInTheRange(sheet: XSSFSheet, columnLater: String, toLater: String, fromLater: String): MutableList<Double> {
    val list = sheet.asSequence()
        .map {
            it.filter { it.address.toString().contains(columnLater) }
                .filter { it.address.toString() > toLater && it.address.toString() < fromLater }
                .filter { it.cellTypeEnum === CellType.NUMERIC }
                .map { it.numericCellValue }
        }
        .filter { it.isNotEmpty() }
        .flatMap { it.toList() }
        .toList()

    return list/*.map { (it * 100).roundToInt() / 100.0 }.map { it.toString() }*/ as MutableList<Double>
}

fun addAverageMaxMin(list: MutableList<Double>): MutableList<String> {
    val aver = (list.average() * 100).roundToInt() / 100.0
    val std = list.map { (it - aver).pow(2.0) }.reduce { acc, it -> it + acc }.div(list.size.toDouble())
    val max = list.maxOrNull()
    val min = list.minOrNull()
    list.add(aver)
    list.add((sqrt(std)))
    list.add(max!!)
    list.add(min!!)
    return list.map { (it * 100).roundToInt() / 100.0 }.map { it.toString() } as MutableList<String>
}

fun addAverageMaxMinSkip(skip: Int, list: MutableList<Double>): MutableList<String> {
    val aver = (list.average() * 100).roundToInt() / 100.0
    val std = list.map { (it - aver).pow(2.0) }.reduce { acc, it -> it + acc }.div(list.size.toDouble())
    val max = list.maxOrNull()
    val min = list.minOrNull()

    val listLocal = (1..skip).toMutableList()
    listLocal.fill(0)
    list.addAll(listLocal.map { it.toDouble() })

    list.add(aver)
    list.add((sqrt(std)))
    list.add(max!!)
    list.add(min!!)
    return list.map { (it * 100).roundToInt() / 100.0 }.map { it.toString() } as MutableList<String>

}

fun cellFull(
    sheet: XSSFSheet,
    rowIndex: Int,
    rowStart: XSSFRow,
    index: Int,
    cellOne: Int,
    biasOne: Int,
    cellTwo: Int,
    biasTwo: Int,
    styleOne: XSSFCellStyle,
    styleTwo: XSSFCellStyle,
    listText: List<String>,
    listValue: List<String>
) {
    val cellText = rowStart.createCell(cellOne)
    val cellValue = rowStart.createCell(cellTwo)
    cellText.setCellValue(listText[index])
    cellValue.setCellValue(listValue[index])

    cellText.cellStyle = styleOne
    cellValue.cellStyle = styleTwo
    if (cellOne != biasOne)
        sheet.addMergedRegion(CellRangeAddress(rowIndex, rowIndex, cellOne, biasOne))
    if (cellTwo != biasTwo)
        sheet.addMergedRegion(CellRangeAddress(rowIndex, rowIndex, cellTwo, biasTwo))
}


//todo createHeaderOne
fun createHeaderOne(
    sheet: XSSFSheet,
    row: XSSFRow,
    index: Int,
    cellOne: Int,
    cellTwo: Int,
    generalStyle: XSSFCellStyle,
    listText: List<String>,
    listValue: List<String>
) {
    val cellHeaderOne1 = row.createCell(cellOne)

    if (cellOne != cellTwo)
        sheet.addMergedRegion(CellRangeAddress(2, 3, cellOne, cellTwo))
    cellHeaderOne1.setCellValue(listText[index] + listValue[index])
    cellHeaderOne1.cellStyle = generalStyle
}


fun main(args: Array<String>) {
    runApplication<ReportLaboratoryMdfApplication>(*args)
}
