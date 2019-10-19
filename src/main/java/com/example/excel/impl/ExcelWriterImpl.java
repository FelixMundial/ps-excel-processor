package com.example.excel.impl;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Calendar;
import java.util.Date;
import java.util.Map;

/**
 * @author yinfelix
 */
public class ExcelWriterImpl {

    private String dateFormat;
    private String doubleFormat;
    private long startTimeMillis;

    private Workbook workbook;
    private Sheet sheet;
    private Row row;

    private static final String DEFAULT_SHEETNAME = "sheet";

    private ExcelReaderImpl excelReader = null;

    private CellStyle newCellStyle = null;

    private static FileOutputStream outStream = null;

    public ExcelWriterImpl(String file, String outputFilePath) {
        try {
            this.excelReader = new ExcelReaderImpl(file);
            generateWorkbook(outputFilePath);
            this.dateFormat = "yyyy-MM-dd HH:mm:ss";
            this.doubleFormat = "0.000000";
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public ExcelWriterImpl(String outputFilePath) {}

    public ExcelWriterImpl() {}

    public ExcelReaderImpl getExcelReader() {
        return excelReader;
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public Sheet getSheet() {
        return sheet;
    }

    private void generateWorkbook(String outputFilePath) throws IOException {
        outStream = new FileOutputStream(outputFilePath);
        workbook = new XSSFWorkbook();

        int i = 1;
        while ((sheet = workbook.getSheet(DEFAULT_SHEETNAME + i)) != null) {
            sheet = workbook.getSheet(DEFAULT_SHEETNAME + ++i);
        }
        sheet = workbook.createSheet(DEFAULT_SHEETNAME + i);

        startTimeMillis = System.currentTimeMillis();
    }

    private void generateSheetHeader(Sheet sheet, int headerRowIndex, Map<Integer, Object> values) {
        row = sheet.createRow(headerRowIndex);

        for (Integer cellIndex : values.keySet()) {
            Cell cell = row.createCell(cellIndex);
            Object value = values.get(cellIndex);
            setValueWithinCell(value, cell);
        }
    }

    /**
     * 在指定Sheet页的指定行列处插入数据，不携带样式信息
     * @param sheet Sheet页名称
     * @param rowIndex 待插入数据所在行号
     * @param columnIndex 待插入数据所在列号
     * @param value 待插入数据的数值
     */
    public void createCellWithValue(Sheet sheet, int rowIndex, int columnIndex, Object value) {
        if ((row = sheet.getRow(rowIndex)) == null) {
            row = sheet.createRow(rowIndex);
        }
        Cell cell = row.createCell(columnIndex);
        setValueWithinCell(value, cell);
    }

    /**
     * 在指定Sheet页的指定行列处插入数据，并携带样式信息
     * @param sheet Sheet页名称
     * @param rowIndex 待插入数据所在行号
     * @param columnIndex 待插入数据所在列号
     * @param value 待插入数据的数值
     * @param targetCellStyle 待生成单元格的样式
     */
    public void createCellWithValue(Sheet sheet, int rowIndex, int columnIndex, Object value, CellStyle targetCellStyle) {
        if (newCellStyle == null) {
            newCellStyle = sheet.getWorkbook().createCellStyle();
        }
        createCellWithValue(sheet, rowIndex, columnIndex, value);

        newCellStyle.cloneStyleFrom(targetCellStyle);
        newCellStyle.setAlignment(targetCellStyle.getAlignment());
        newCellStyle.setFillForegroundColor(targetCellStyle.getFillForegroundColor());
        newCellStyle.setFillBackgroundColor(targetCellStyle.getFillBackgroundColor());
        newCellStyle.setFillPattern(targetCellStyle.getFillPattern());
        newCellStyle.setDataFormat(targetCellStyle.getDataFormat());
        newCellStyle.setHidden(targetCellStyle.getHidden());
        newCellStyle.setIndention(targetCellStyle.getIndention());
        newCellStyle.setLocked(targetCellStyle.getLocked());
        newCellStyle.setRotation(targetCellStyle.getRotation());
        newCellStyle.setVerticalAlignment(targetCellStyle.getVerticalAlignment());
        newCellStyle.setWrapText(targetCellStyle.getWrapText());

        newCellStyle.setBorderBottom(targetCellStyle.getBorderBottom());
        newCellStyle.setBorderLeft(targetCellStyle.getBorderLeft());
        newCellStyle.setBorderRight(targetCellStyle.getBorderRight());
        newCellStyle.setBorderTop(targetCellStyle.getBorderTop());
        newCellStyle.setBottomBorderColor(targetCellStyle.getBottomBorderColor());
        newCellStyle.setLeftBorderColor(targetCellStyle.getLeftBorderColor());
        newCellStyle.setRightBorderColor(targetCellStyle.getRightBorderColor());
        newCellStyle.setTopBorderColor(targetCellStyle.getTopBorderColor());
        sheet.getRow(rowIndex).getCell(columnIndex).setCellStyle(newCellStyle);
    }

    private void setValueWithinCell(Object value, Cell cell) {
        if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value instanceof Double) {
            cell.setCellValue((Double) value);
        } else if (value instanceof Date) {
            cell.setCellValue((Date) value);
        } else if (value instanceof Calendar) {
            cell.setCellValue((Calendar) value);
        } else if (value instanceof RichTextString) {
            cell.setCellValue((RichTextString) value);
        }
    }

    /**
     * 将工作簿写入输出文件
     * @throws IOException IO操作
     */
    public void commitWorkbook() throws IOException {
        if (workbook != null) {
            workbook.write(outStream);
        }
    }

    /**
     * 关闭输出流成员
     * @throws IOException IO操作
     */
    public void close() throws IOException {
        outStream.close();
        System.out.println("文件生成时间：" + (System.currentTimeMillis() - this.startTimeMillis) + "毫秒，约等于" + String.format("%.1f", (System.currentTimeMillis() - this.startTimeMillis) / 1000.0f) + "秒");
    }

    public static void main(String[] args) {}
}