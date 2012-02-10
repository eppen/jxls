package com.jxls.writer.transform.poi

import spock.lang.Specification
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Row

import com.jxls.writer.command.Context
import com.jxls.writer.CellData
import com.jxls.writer.Pos
import org.apache.poi.ss.util.CellRangeAddress

/**
 * @author Leonid Vysochyn
 * Date: 1/23/12 3:23 PM
 */
class PoiTransformerTest extends Specification{
    Workbook wb;

    def setup(){
        wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("sheet 1")
        Row row0 = sheet.createRow(0)
        row0.createCell(0).setCellValue(1.5)
        row0.createCell(1).setCellValue('${x}')
        row0.createCell(2).setCellValue('${x*y}')
        row0.createCell(3).setCellValue('Merged value')
        sheet.addMergedRegion(new CellRangeAddress(0,1,3,4))
        row0.setHeight((short)23)
        sheet.setColumnWidth(1, 123);
        Row row1 = sheet.createRow(1)
        row1.createCell(1).setCellFormula("SUM(A1:A3)")
        row1.createCell(2).setCellValue('${y*y}')
        row1.setHeight((short)456)
        Row row2 = sheet.createRow(2)
        row2.createCell(0).setCellValue("XYZ")
        row2.createCell(1).setCellValue('${2*y}')
        row2.createCell(2).setCellValue('${4*4}')
        row2.createCell(3).setCellValue('${2*x}x and ${2*y}y')
        row2.createCell(4).setCellValue('$[${myvar}*SUM(A1:A5) + ${myvar2}]')
    }

    def "test template cells storage"(){
        when:
            def poiTransformer = PoiTransformer.createTransformer(wb)
            wb.removeSheetAt(0)
        then:
            assert wb.getNumberOfSheets() == 0
            assert poiTransformer.getCellData(new Pos("sheet 1", row, col)).getCellValue() == value
        where:
            row | col   | value
            0   | 0     | new Double(1.5)
            0   | 1     | '${x}'
            0   | 2     | '${x*y}'
            1   | 1     | "SUM(A1:A3)"
            2   | 0     | "XYZ"
            2   | 4     |  '$[${myvar}*SUM(A1:A5) + ${myvar2}]'
    }

    def "test transform string var"(){
        given:
            def poiTransformer = PoiTransformer.createTransformer(wb)
            def context = new Context()
            context.putVar("x", "Abcde")
        when:
            poiTransformer.transform(new Pos("sheet 1",0, 1), new Pos("sheet 1",7, 7), context)
        then:
            Sheet sheet = wb.getSheetAt(0)
            Row row7 = sheet.getRow(7)
            row7.getCell(7).getStringCellValue() == "Abcde"
            sheet.getColumnWidth(7) == 123
            sheet.getRow(7).getHeight() == 23
    }

    def "test transform numeric var"(){
        given:
            def poiTransformer = PoiTransformer.createTransformer(wb)
            def context = new Context()
            context.putVar("x", 3)
            context.putVar("y", 5)
        when:
            poiTransformer.transform(new Pos("sheet 1",0, 2), new Pos("sheet 2",7, 7), context)
        then:
            Sheet sheet = wb.getSheet("sheet 2")
            Row row7 = sheet.getRow(7)
            row7.getCell(7).cellType == org.apache.poi.ss.usermodel.Cell.CELL_TYPE_NUMERIC
            row7.getCell(7).getNumericCellValue() == 15
    }

    def "test transform formula cell"(){
        given:
            def poiTransformer = PoiTransformer.createTransformer(wb)
            def context = new Context()
        when:
            poiTransformer.transform(new Pos("sheet 1",1, 1), new Pos("sheet 2",7, 7), context)
        then:
            Sheet sheet = wb.getSheet("sheet 2")
            Row row7 = sheet.getRow(7)
            row7.getCell(7).cellType == org.apache.poi.ss.usermodel.Cell.CELL_TYPE_FORMULA
            row7.getCell(7).getCellFormula() == "SUM(A1:A3)"
    }

    def "test transform a cell to other sheet"(){
        given:
            def poiTransformer = PoiTransformer.createTransformer(wb)
            def context = new Context()
            context.putVar("x", "Abcde")
        when:
            poiTransformer.transform(new Pos("sheet 1",0, 1), new Pos("sheet2", 7, 7), context)
        then:
            Sheet sheet = wb.getSheet("sheet 1")
            Row row = sheet.getRow(7)
        //TODO
            row == null
            Sheet sheet1 = wb.getSheet("sheet2")
            Row row1 = sheet1.getRow(7)
            row1.getCell(7).getStringCellValue() == "Abcde"
    }
    
    def "test transform multiple times"(){
        given:
            def poiTransformer = PoiTransformer.createTransformer(wb)
            def context1 = new Context()
            context1.putVar("x", "Abcde")
            def context2 = new Context()
            context2.putVar("x", "Fghij")
        when:
            poiTransformer.transform(new Pos("sheet 1",0, 1), new Pos("sheet 1",5, 1), context1)
            poiTransformer.transform(new Pos("sheet 1",0, 1), new Pos("sheet 2",10, 1), context2)
        then:
            Sheet sheet = wb.getSheet("sheet 1")
            Sheet sheet2 = wb.getSheet("sheet 2")
            Row row5 = sheet.getRow(5)
            Row row10 = sheet2.getRow(10)
            row5.getCell(1).getStringCellValue() == "Abcde"
            row10.getCell(1).getStringCellValue() == "Fghij"
    }

    def "test transform overridden cells"(){
        given:
        def poiTransformer = PoiTransformer.createTransformer(wb)
        def context1 = new Context()
        context1.putVar("x", "Abcde")
        def context2 = new Context()
        context2.putVar("x", "Fghij")
        when:
        poiTransformer.transform(new Pos("sheet 1",0, 1), new Pos("sheet 1",5, 1), context1)
        poiTransformer.transform(new Pos("sheet 1",0, 0), new Pos("sheet 2",0, 1), context1)
        poiTransformer.transform(new Pos("sheet 1",0, 1), new Pos("sheet 2",10, 1), context2)
        then:
        Sheet sheet = wb.getSheet("sheet 1")
        Row row5 = sheet.getRow(5)
        row5.getCell(1).getStringCellValue() == "Abcde"
        Sheet sheet2 = wb.getSheet("sheet 2")
        sheet2.getRow(0).getCell(1).getNumericCellValue() == 1.5
        Row row10 = sheet2.getRow(10)
        row10.getCell(1).getCellType() == org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING
        row10.getCell(1).getStringCellValue() == "Fghij"
    }

    def "test multiple expressions in a cell"(){
        given:
            def poiTransformer = PoiTransformer.createTransformer(wb)
            def context = new Context()
            context.putVar("x", 2)
            context.putVar("y", 3)
        when:
            poiTransformer.transform(new Pos("sheet 1",2, 3), new Pos("sheet 2",7, 7), context)
        then:
            Sheet sheet = wb.getSheet("sheet 2")
            Row row7 = sheet.getRow(7)
            row7.getCell(7).getStringCellValue() == "4x and 6y"
    }

    def "test ignore source column and row props"(){
        given:
            def poiTransformer = PoiTransformer.createTransformer(wb)
            poiTransformer.setIgnoreColumnProps(true)
            poiTransformer.setIgnoreRowProps(true)
            def context = new Context()
            context.putVar("x", "Abcde")
        when:
            poiTransformer.transform(new Pos("sheet 1",0, 1), new Pos("sheet 2",7, 7), context)
        then:
            Sheet sheet1 = wb.getSheet("sheet 1")
            Sheet sheet2 = wb.getSheet("sheet 2")
            sheet2.getColumnWidth(7) != sheet1.getColumnWidth(1)
            sheet1.getRow(0).getHeight() != sheet2.getRow(7).getHeight()
    }

    def "test set formula value"(){
        given:
            def poiTransformer = PoiTransformer.createTransformer(wb)
        when:
            poiTransformer.setFormula(new Pos("sheet 2",1, 1), "SUM(B1:B5)")
        then:
            wb.getSheet("sheet 2").getRow(1).getCell(1).getCellFormula() == "SUM(B1:B5)"
    }

    def "test get formula cells"(){
        when:
            def poiTransformer = PoiTransformer.createTransformer(wb)
            def context = new Context()
            context.putVar("myvar", 2)
            context.putVar("myvar2", 4)
            poiTransformer.transform(new Pos("sheet 1",2,4), new Pos("sheet 2", 10,10), context)
            def formulaCells = poiTransformer.getFormulaCells()
        then:
            formulaCells.size() == 2
        formulaCells.contains(new CellData("sheet 1",1,1, CellData.CellType.FORMULA, "SUM(A1:A3)"))
            formulaCells.contains(new CellData("sheet 1",2,4, CellData.CellType.STRING, '$[${myvar}*SUM(A1:A5) + ${myvar2}]'))
    }

    def "test get target cells"(){
        when:
            def poiTransformer = PoiTransformer.createTransformer(wb)
            def context = new Context()
            poiTransformer.transform(new Pos("sheet 1",1,1), new Pos("sheet 2",10,10), context)
            poiTransformer.transform(new Pos("sheet 1",1,1), new Pos("sheet 1",10,12), context)
            poiTransformer.transform(new Pos("sheet 1",1,1), new Pos("sheet 1",10,14), context)
            poiTransformer.transform(new Pos("sheet 1",2,1), new Pos("sheet 2",20,11), context)
            poiTransformer.transform(new Pos("sheet 1",2,1), new Pos("sheet 1",20,12), context)
        then:
            poiTransformer.getTargetPos(new Pos("sheet 1",1,1)).toArray() == [new Pos("sheet 2",10,10), new Pos("sheet 1",10,12), new Pos("sheet 1",10,14)]
            poiTransformer.getTargetPos(new Pos("sheet 1",2,1)).toArray() == [new Pos("sheet 2",20,11), new Pos("sheet 1",20,12)]
    }

    def "test reset target cells"(){
        when:
            def poiTransformer = PoiTransformer.createTransformer(wb)
            def context = new Context()
            poiTransformer.transform(new Pos("sheet 1",1,1), new Pos("sheet 1",10,10), context)
            poiTransformer.transform(new Pos("sheet 1",2,1), new Pos("sheet 1",20,11), context)
            poiTransformer.resetTargetCells()
            poiTransformer.transform(new Pos("sheet 1",1,1), new Pos("sheet 2",10,12), context)
            poiTransformer.transform(new Pos("sheet 1",1,1), new Pos("sheet 1",10,14), context)
            poiTransformer.transform(new Pos("sheet 1",2,1), new Pos("sheet 1",20,12), context)
        then:
            poiTransformer.getTargetPos(new Pos("sheet 1",1,1)).toArray() == [new Pos("sheet 2",10,12), new Pos("sheet 1",10,14)]
            poiTransformer.getTargetPos(new Pos("sheet 1",2,1)).toArray() == [new Pos("sheet 1",20,12)]
    }

    def "test transform merged cells"(){
        when:
            def poiTransformer = PoiTransformer.createTransformer(wb)
            def context = new Context()
            wb.getSheetAt(0).getNumMergedRegions() == 1
            poiTransformer.transform(new Pos("sheet 1",0,3), new Pos("sheet 1",10,10), context)
        then:
            Sheet sheet = wb.getSheet("sheet 1")
            sheet.getNumMergedRegions() == 2
            sheet.getMergedRegion(1).toString() == new CellRangeAddress(10,11,10,11).toString()
    }

}
