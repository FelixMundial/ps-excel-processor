package com.example.excel.impl;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

/**
 * @apiNote 报表拆分工具类：指定列号并按列对报表进行分类拆分（携带样式信息）
 * @author yinfelix
 */
public class ExcelSplitByRowProcessorImpl {

    private String destFile;
    private String sheetName;
    private int rowStart;

    private ExcelUtils utils;
    private ExcelReaderImpl excelReader = null;
    private ExcelWriterImpl excelWriter = null;

    public ExcelSplitByRowProcessorImpl(String sourceFile, String destFile, String sheetName, int rowStart) {
        this.destFile = destFile;
        this.sheetName = sheetName;
        this.rowStart = rowStart - 1;

        this.utils = new ExcelUtils();
        this.excelWriter = new ExcelWriterImpl(sourceFile, destFile);
        this.excelReader = excelWriter.getExcelReader();
    }

    public ExcelSplitByRowProcessorImpl() {
    }

    public ExcelWriterImpl getExcelWriter() {
        return excelWriter;
    }

    public ExcelReaderImpl getExcelReader() {
        return excelReader;
    }

    /**
     * 报表拆分入口方法（列号为字母）
     * @param splitCondition 拆分条件
     * @param targetColumnLabel 拆分列号（字母）
     */
    public void doExcelRowSplit(String splitCondition, String targetColumnLabel) {
        doExcelRowSplit(splitCondition, utils.getColIndexFromColLabel(targetColumnLabel) + 1);
    }

    /**
     * 报表拆分入口方法（列号为数字）
     * @param splitCondition 拆分条件
     * @param targetColumnIndex 拆分列号（数字）
     */
    public void doExcelRowSplit(String splitCondition, int targetColumnIndex) {
        int sheetNumber = excelReader.getSheetIndexFromName(sheetName);

        try {
            Workbook currentWorkbook = excelReader.getWorkbook();
            Sheet currentSheet = currentWorkbook.getSheetAt(sheetNumber - 1);

            int sourceRowCount = excelReader.getRowCount(sheetNumber) - 1;
//            int sourceColCount = excelReader.getColCount(sheetNumber, sourceRowCount);
            int tempRowCount = sourceRowCount;

            Map<Integer, Map<Integer, String>> formula = getFormula(currentWorkbook, currentSheet, rowStart, tempRowCount);
//            Map<Integer, Map<Integer, Comment>> comment = getComment(currentWorkbook, currentSheet, rowStart, tempRowCount);
            Map<Integer, Map<Integer, Comment>> comment = getComment(currentWorkbook, currentSheet, 0, tempRowCount);

//            从后向前剔除数据区域中特定行
            for (int rowIndex = sourceRowCount; rowIndex >= rowStart; rowIndex--) {
                if (null != currentSheet.getRow(rowIndex)) {
                    if (!splitCondition.equals(excelReader.getCellValue(sheetNumber, rowIndex + 1, targetColumnIndex))) {
                        removeRow(currentSheet, rowIndex, tempRowCount--);
                    }
                }
            }

            for (int rowIndex = rowStart; rowIndex < tempRowCount; rowIndex++) {
                for (Integer columnIndex : formula.get(rowIndex).keySet()) {
                    currentSheet.getRow(rowIndex).getCell(columnIndex).setCellFormula(formula.get(rowIndex).get(columnIndex));
                }
            }

            for (int rowIndex = 0; rowIndex < tempRowCount; rowIndex++) {
                for (Integer columnIndex : comment.get(rowIndex).keySet()) {
                    currentSheet.getRow(rowIndex).getCell(columnIndex).setCellComment(comment.get(rowIndex).get(columnIndex));
                }
                updateFormula(currentWorkbook, currentSheet, rowIndex);
            }
//            currentWorkbook.setForceFormulaRecalculation(true);

            excelReader.exportWorkbook(destFile);
            excelReader.close();
        } catch (IOException e) {}
    }

    /**
     * 剔除报表中的特定行
     * @param sheet 待剔除行所在的sheet页
     * @param rowIndex 待剔除行所在的行号
     * @param rowEnd 待剔除行所在sheet页的末行，通过进行自减的参数传入，而非每次反复获取
     */
    private void removeRow(Sheet sheet, int rowIndex, int rowEnd) {
        if (rowIndex >= 0 && rowIndex < rowEnd) {
            sheet.removeRow(sheet.getRow(rowIndex));
            // 将行号为rowIndex+1一直到行号为lastRowNum的单元格全部上移一行，以便删除rowIndex行
            sheet.shiftRows(rowIndex + 1, rowEnd, -1);
        }
        if (rowIndex == rowEnd) {
            Row removingRow = sheet.getRow(rowIndex);
            if (removingRow != null) {
                sheet.removeRow(removingRow);
            }
        }
    }

    /**
     * 提取报表中所有公式字符串
     * @param workbook 待提取公式所在workbook
     * @param sheet 待提取公式所在sheet页
     * @param rowStart 待提取公式区域的首行索引
     * @param rowEnd 待提取公式区域的末行索引
     * @return 包含每行公式信息Map的二维Map，K-V分别为行号和对应行公式信息Map；其中每行公式信息Map的K-V分别为列号和对应列的公式字符串
     */
    private Map<Integer, Map<Integer, String>> getFormula(Workbook workbook, Sheet sheet, int rowStart, int rowEnd) {
        Map<Integer, Map<Integer, String>> formulaMap = new HashMap<Integer, Map<Integer, String>>();
        HashMap<Integer, String> formulaMapInsideRow;
        for (int rowIndex = rowStart; rowIndex <= rowEnd; rowIndex++) {
            formulaMapInsideRow = new HashMap<Integer, String>();
            Row row = sheet.getRow(rowIndex);
            Cell cell;
            CellStyle cellStyle;
            if (null != row) {
                for (int columnIndex = row.getFirstCellNum(); columnIndex < row.getLastCellNum(); columnIndex++) {
                    cell = row.getCell(columnIndex);
                    if (null != cell && cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
                        formulaMapInsideRow.put(columnIndex, cell.getCellFormula());
                        cellStyle = cell.getCellStyle();
                        row.removeCell(cell);
                        row.createCell(columnIndex);
                        row.getCell(columnIndex).setCellStyle(cellStyle);
                    }
                }
                formulaMap.put(rowIndex, formulaMapInsideRow);
            }
        }
        return formulaMap;
    }

    private Map<Integer, Map<Integer, Comment>> getComment(Workbook workbook, Sheet sheet, int rowStart, int rowEnd) {
        Map<Integer, Map<Integer, Comment>> commentMap = new HashMap<Integer, Map<Integer, Comment>>();
        HashMap<Integer, Comment> commentMapInsideRow = null;
        for (int rowIndex = rowStart; rowIndex <= rowEnd; rowIndex++) {
            commentMapInsideRow = new HashMap<Integer, Comment>();
            Row row = sheet.getRow(rowIndex);
            Cell cell;
            if (null != row) {
                for (int columnIndex = row.getFirstCellNum(); columnIndex < row.getLastCellNum(); columnIndex++) {
                    cell = row.getCell(columnIndex);
                    if (null != cell && null != cell.getCellComment()) {
                        commentMapInsideRow.put(columnIndex, cell.getCellComment());
                    }
                }
                commentMap.put(rowIndex, commentMapInsideRow);
            }
        }
        return commentMap;
    }

    /**
     * 强制更新报表中所有公式
     * @param workbook 公式所在workbook
     * @param sheet 公式所在sheet页
     * @param rowIndex 公式所在行号
     */
    private void updateFormula(Workbook workbook, Sheet sheet, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        Cell cell;
        FormulaEvaluator eval = new XSSFFormulaEvaluator((XSSFWorkbook) workbook);
        if (null != row) {
            for (int columnIndex = row.getFirstCellNum(); columnIndex < row.getLastCellNum(); columnIndex++) {
                cell = row.getCell(columnIndex);
                if (null != cell && cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
                    eval.evaluateFormulaCell(cell);
                }
            }
        }
    }

    public static void main(String[] args) {}
}
