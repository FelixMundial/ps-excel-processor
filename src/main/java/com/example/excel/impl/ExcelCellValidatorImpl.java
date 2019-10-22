package com.example.excel.impl;

import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddressList;

import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author yinfelix
 */
public class ExcelCellValidatorImpl {

    private ExcelUtils utils;
    private ExcelReaderImpl excelReader;
    private ExcelCellValidatorImpl excelCellValidator;

    private String destFile;
    private String sheetName;
    private String colIndex;
    private int thresholdValue;
    private String validationFormula;
    private Map<String, Integer> errorStyleMap;

    ExcelCellValidatorImpl() {
    }

    public ExcelCellValidatorImpl(String sourceFile, String destFile, String sheetName) {
        this.utils = new ExcelUtils();
        this.excelReader = new ExcelReaderImpl(sourceFile);
        this.excelCellValidator = new ExcelCellValidatorImpl();
        this.sheetName = sheetName;
        this.destFile = destFile;

        errorStyleMap = new HashMap<>();
//        定制化数据验证警告样式映射
        errorStyleMap.put("100", DataValidation.ErrorStyle.INFO);
        errorStyleMap.put("101", DataValidation.ErrorStyle.WARNING);
        errorStyleMap.put("102", DataValidation.ErrorStyle.STOP);
    }

    /**
     * 总额控制入口方法（在指定Sheet页指定列内插入数据验证规则）
     * @param rowStart 总额控制圈注区域行首
     * @param colLabel 总额控制列号（字母列号）
     * @param validationFormula 总额控制公式
     * @param alertStyle 出错警告样式
     * @param alertTitle 出错警告标题
     * @param alertContent 出错警告错误信息
     */
    public void doAddCellValidation(int rowStart, String colLabel, String validationFormula, String alertStyle, String alertTitle, String alertContent) {
        try {
            Sheet currentSheet = excelReader.getWorkbook().getSheet(sheetName);
            int colIndex = utils.getColIndexFromColLabel(colLabel.toUpperCase());

            DataValidation dataValidation = excelCellValidator
                    .getRegionalDataValidationWithCustomFormula(currentSheet, (short) (rowStart - 1), (short) colIndex, (short) currentSheet.getLastRowNum(), (short) colIndex, validationFormula);

            if (dataValidation != null) {
                excelCellValidator.setDataValidationStyle
                        (dataValidation, false, errorStyleMap.get(alertStyle), alertTitle, alertContent);
                currentSheet.addValidationData(dataValidation);
            }
            excelReader.exportWorkbook(this.destFile);
            excelReader.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 总额控制入口方法，提供默认警告样式
     * @param rowStart 总额控制圈注区域行首
     * @param colLabel 总额控制列号（字母列号）
     * @param formula 总额控制公式
     */
    public void doAddCellValidation(int rowStart, String colLabel, String formula) {
        String alertStyle = "100";
        String alertTitle = "数据验证失败！";
        String alertContent = "数据验证失败！";
        doAddCellValidation(rowStart, colLabel, formula, alertStyle, alertTitle, alertContent);
    }

    /**
     * 总额控制入口方法（在指定Sheet页指定列内插入数据验证规则），此处公式为硬编码对业务定制化格式
     * @param rowStart 总额控制圈注区域行首
     * @param colLabel 总额控制列号（字母列号）
     * @param thresholdValue 总额控制阈值
     * @param alertStyle 出错警告样式
     * @param alertTitle 出错警告标题
     * @param alertContent 出错警告错误信息
     */
    public void doAddCellValidation(int rowStart, String colLabel, int thresholdValue, String alertStyle, String alertTitle, String alertContent) {
        String formula = "=SUM($" + colLabel  + ":$" + colLabel + ") <=  " + thresholdValue;
        doAddCellValidation(rowStart, colLabel, formula, alertStyle, alertTitle, alertContent);
    }

    /**
     * 获取跨单元格数据验证对象
     * @param sheet 数据验证所在区域sheet页
     * @param firstRowIndex 数据验证所在区域行首行号（从0开始）
     * @param firstColIndex 数据验证所在区域列首列号（从0开始）
     * @param endRowIndex 数据验证所在区域行尾行号（从0开始）
     * @param endColIndex 数据验证所在区域列尾列号（从0开始）
     * @param formula 数据验证公式（行列号从0开始）
     * @return 跨单元格数据验证对象
     */
    DataValidation getRegionalDataValidationWithCustomFormula(Sheet sheet, short firstRowIndex, short firstColIndex, short endRowIndex, short endColIndex, String formula) {
        DataValidationHelper helper = sheet.getDataValidationHelper();
        DataValidationConstraint constraint = helper.createCustomConstraint(formula);
        CellRangeAddressList addressList = new CellRangeAddressList(firstRowIndex, endRowIndex, firstColIndex, endColIndex);
        System.out.println("已添加数据校验");

        return helper.createValidation(constraint, addressList);
    }

    /**
     * 设置跨单元格数据验证样式
     * @param validation 跨单元格数据验证对象
     * @param ignoresEmptyValue 是否忽略空值
     * @param alertStyle 出错警告样式
     * @param alertTitle 出错警告标题
     * @param alertContent 出错警告错误信息
     * @see DataValidation.ErrorStyle
     */
    void setDataValidationStyle(DataValidation validation, boolean ignoresEmptyValue, int alertStyle, String alertTitle, String alertContent) {
        validation.setEmptyCellAllowed(!ignoresEmptyValue);
        // 整合getDataValidationList()
        validation.setSuppressDropDownArrow(true);
        // 选定单元格时是否显示输入信息
        validation.setShowPromptBox(true);
        validation.createPromptBox("", "");
        // 输入无效数据时是否显示出错警告
        validation.setShowErrorBox(true);
        validation.setErrorStyle(alertStyle);
        validation.createErrorBox(alertTitle, alertContent);
        System.out.println("已设置" + alertStyle + "数据校验样式");
    }

    private DataValidation getDataValidationList(Sheet sheet, short firstRowIndex, short firstColIndex, short endRowIndex, short endColIndex, List<String> strList) {
        String[] data = strList.toArray(new String[0]);
        DataValidationHelper helper = sheet.getDataValidationHelper();
        DataValidationConstraint constraint = helper.createExplicitListConstraint(data);
        CellRangeAddressList addressList = new CellRangeAddressList(firstRowIndex, endRowIndex, firstColIndex, endColIndex);

        return helper.createValidation(constraint, addressList);
    }

    /**
     * 在指定Sheet页指定区域内插入数据验证规则
     * @param rowStart 指定区域行首
     * @param rowEnd 指定区域行尾
     * @param colStart 指定区域列首
     * @param colEnd 指定区域列尾
     * @param formula 数据验证规则公式
     * @param alertStyle 出错警告样式
     * @param alertTitle 出错警告标题
     * @param alertContent 出错警告错误信息
     */
    public void addCellValidation(int rowStart, int rowEnd, int colStart, int colEnd, String formula, String alertStyle, String alertTitle, String alertContent) {
        try {
            Sheet currentSheet = excelReader.getWorkbook().getSheet(sheetName);
            DataValidation dataValidation = excelCellValidator
                    .getRegionalDataValidationWithCustomFormula(currentSheet, (short) rowStart, (short) colStart, (short) rowEnd, (short) colEnd, formula);

            if (dataValidation != null) {
                excelCellValidator.setDataValidationStyle
                        (dataValidation, false, errorStyleMap.get(alertStyle), alertTitle, alertContent);
                currentSheet.addValidationData(dataValidation);
            }
            excelReader.exportWorkbook(this.destFile);
            excelReader.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {}
}
