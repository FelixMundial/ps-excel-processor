package com.example.excel.impl;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * @author yinfelix
 */
public class ExcelReaderImpl {
    private String dateFormat;
    private String doubleFormat;
    private long startTimeMillis;

    private Workbook workbook;
    private Sheet sheet;
    private Row row;

    private static final String XSSF_SUFFIX = ".xlsx";
    private static final String DOT_STRING = ".";

    private static InputStream inStream = null;
    private static int colCountResult = 0;

    public ExcelReaderImpl(String inputFile) {
        try {
            this.loadFileAsWorkbook(inputFile);
            this.dateFormat = "yyyy-MM-dd HH:mm:ss";
            this.doubleFormat = "0.000000";
        } catch (IOException e) {}
    }

    public ExcelReaderImpl() {}

    public Workbook getWorkbook() {
        return workbook;
    }

    /**
     * 通过指定文件初始化工作簿
     * @param file 报表文件路径
     * @throws IOException IO操作
     */
    public void loadFileAsWorkbook(String file) throws IOException {
        inStream = new FileInputStream(file);
        if (file.endsWith(XSSF_SUFFIX)) {
            this.workbook = new XSSFWorkbook(inStream);
            this.startTimeMillis = System.currentTimeMillis();
        } else {
            throw new RuntimeException("文件格式错误！");
        }
    }

    /**
     * 获取工作簿sheet页数
     * @return sheet页数
     */
    public int getNumberOfSheets() {
        return getWorkbook().getNumberOfSheets();
    }

    /**
     * 通过sheet页ID（从1开始）获取sheet页名称
     * @param sheetIndex sheet页ID（从1开始）
     * @return sheet页名称
     */
    public String getSheetName(int sheetIndex) {
        --sheetIndex;
        return getWorkbook().getSheetName(sheetIndex);
    }

    /**
     * 通过sheet页名称获取sheet页ID（从1开始）
     * @param sheetName sheet页名称
     * @return sheet页ID（从1开始）
     */
    public int getSheetIndexFromName(String sheetName) {
        for (int sheetIndex = 1; sheetIndex <= getNumberOfSheets(); sheetIndex++) {
            if (getSheetName(sheetIndex).equals(sheetName)) {
                return sheetIndex;
            }
        }
        return -1;
    }

    /**
     * 获取指定sheet页的有效行数
     * @param sheetIndex sheet页ID（从1开始）
     * @return 指定sheet页的有效行数
     */
    public int getRowCount(int sheetIndex) {
        --sheetIndex;
        int nullRow = 0;
        this.sheet = getWorkbook().getSheetAt(sheetIndex);
        int firstRowIndex = this.sheet.getFirstRowNum();
        int lastRowIndex = this.sheet.getLastRowNum();
        int rowCount = firstRowIndex + 1;

        for(int i = firstRowIndex; i <= lastRowIndex; ++i) {
            if (this.sheet.getRow(i) == null) {
                ++nullRow;
            } else {
                if (nullRow != 0) {
                    rowCount += nullRow;
                }

                nullRow = 0;
                ++rowCount;
            }
        }
        return rowCount;
    }

    /**
     * 获取指定sheet页、指定行的有效列数
     * @param sheetIndex sheet页ID（从1开始）
     * @param rowIndex 行号（从1开始）
     * @return 指定sheet页指定行的有效列数
     */
    public int getColCount(int sheetIndex, int rowIndex) {
        if (colCountResult == 0) {
            --sheetIndex;
            this.sheet = getWorkbook().getSheetAt(sheetIndex);
            this.row = this.sheet.getRow(rowIndex);
            while (this.row != null) {
                colCountResult = this.row.getPhysicalNumberOfCells();
                break;
            }
        }
        return colCountResult;
    }

    /**
     * 获取指定sheet页、指定行和指定列的单元格对象
     * @param sheetIndex sheet页ID（从1开始）
     * @param rowIndex 行号（从1开始）
     * @param columnIndex 列号（从1开始）
     * @return 指定位置的单元格对象
     */
    public Cell getCell(int sheetIndex, int rowIndex, int columnIndex) {
        --sheetIndex;
        --rowIndex;
        --columnIndex;
        this.sheet = getWorkbook().getSheetAt(sheetIndex);
        if (this.sheet == null) {
            return null;
        } else {
            this.row = this.sheet.getRow(rowIndex);
            if (this.row != null) {
                return this.row.getCell(columnIndex);
            } else {
                return null;
            }
        }
    }

    /**
     * 获取指定sheet页、指定行和指定列的单元格对象数值
     * @param sheetIndex sheet页ID（从1开始）
     * @param rowIndex 行号（从1开始）
     * @param columnIndex 列号（从1开始）
     * @return 指定位置的单元格对象数值
     */
    public String getCellValue(int sheetIndex, int rowIndex, int columnIndex) {
        --sheetIndex;
        --rowIndex;
        --columnIndex;
        this.sheet = getWorkbook().getSheetAt(sheetIndex);
        if (this.sheet == null) {
            return "";
        } else {
            this.row = this.sheet.getRow(rowIndex);
            if (this.row == null) {
                return "";
            } else {
                Cell cell = this.row.getCell(columnIndex);
                if (cell == null) {
                    return "";
                } else {
                    String value = "";
                    switch(cell.getCellType()) {
                        case Cell.CELL_TYPE_NUMERIC:
                            if (HSSFDateUtil.isCellDateFormatted(cell)) {
                                SimpleDateFormat d = new SimpleDateFormat(this.dateFormat);
                                Date dt = HSSFDateUtil.getJavaDate(cell.getNumericCellValue());
                                value = d.format(dt);
                            } else {
                                double d1 = cell.getNumericCellValue();
                                DecimalFormat f = new DecimalFormat(this.doubleFormat);
                                value = this.subZeroAndDot(f.format(d1));
                            }
                            break;
                        case Cell.CELL_TYPE_STRING:
                            value = cell.getStringCellValue();
                            break;
                        case Cell.CELL_TYPE_FORMULA:
                            value = cell.getCellFormula();
                            break;
                        case Cell.CELL_TYPE_BLANK:
                            value = " ";
                            break;
                        case Cell.CELL_TYPE_BOOLEAN:
                            value = String.valueOf(cell.getBooleanCellValue());
                            break;
                        case Cell.CELL_TYPE_ERROR:
                            value = "";
                        default:break;
                    }
                    return value;
                }
            }
        }
    }

    /**
     * 剔除字符串内的零和小数点
     * @param s 待处理的字符串
     * @return 处理过后的字符串
     */
    public String subZeroAndDot(String s) {
        if (s.indexOf(DOT_STRING) > 0) {
            s = s.replaceAll("0+?$", "");
            s = s.replaceAll("[.]$", "");
        }
        return s;
    }

    /**
     * 获取指定sheet页、指定行和指定列的单元格样式
     * @param sheetIndex sheet页ID（从1开始）
     * @param rowIndex 行号（从1开始）
     * @param columnIndex 列号（从1开始）
     * @return 指定sheet页、指定行和指定列的单元格样式
     */
    public CellStyle getCellStyle(int sheetIndex, int rowIndex, int columnIndex) {
        --sheetIndex;
        --rowIndex;
        --columnIndex;
        return getWorkbook().getSheetAt(sheetIndex).getRow(rowIndex).getCell(columnIndex).getCellStyle();
    }

    /**
     * 将工作簿另存为新文件
     * @param outputFilePath 新文件所在路径
     * @throws IOException IO操作
     */
    public void exportWorkbook(String outputFilePath) throws IOException {
        FileOutputStream outStream = new FileOutputStream(outputFilePath);
        getWorkbook().write(outStream);
    }

    /**
     * 关闭输出流成员
     * @throws IOException IO操作
     */
    public void close() throws IOException {
        inStream.close();
        System.out.println("文件读取时间：" + (System.currentTimeMillis() - this.startTimeMillis) + "毫秒，约等于" + String.format("%.1f", (System.currentTimeMillis() - this.startTimeMillis) / 1000.0f) + "秒");
    }

    public static void main(String[] args) throws IOException {}
}