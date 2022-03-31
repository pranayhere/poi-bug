package com.rateDemo;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
    public static void main(String[] args) {
        double nper = 360.0;
        double pmt = 6.56;
        double pv = -2000.0;

        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();

        XSSFRow row = (XSSFRow) sheet.createRow(1);
        XSSFCell cell = row.createCell(1);

        cell.setCellType(CellType.NUMERIC);
        cell.setCellFormula("RATE(" + nper + ", " + pmt + ", " + pv + ")");
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

        evaluator.evaluateInCell(cell);
        double rate = cell.getNumericCellValue();

        System.out.println("Rate : " + rate);
    }
}
