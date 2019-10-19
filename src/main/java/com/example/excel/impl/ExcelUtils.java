package com.example.excel.impl;

/**
 * @author yinfelix
 */
public class ExcelUtils {

    private static final int ALPHABET_LENGTH = 26;

    /**
     * 由字母列号获取数字列号
     * @param colLabel 列号（字母）
     * @return 数字列号
     */
    public int getColIndexFromColLabel(String colLabel) {
        int resultColIndex;
        int length = colLabel.length();
        if (length == 1) {
            resultColIndex = colLabel.charAt(0) - 'A';
        } else {
            resultColIndex = (getColIndexFromColLabel(colLabel.substring(0, length - 1)) + 1) * 26 + getColIndexFromColLabel(colLabel.substring(length - 1));
        }
        return resultColIndex;
    }

    /**
     * 由数字列号获取字母列号
     * @param colIndex 数字列号
     * @return 字母列号
     */
    public String getColLabelFromColIndex(int colIndex) {
        String resultColLabel = "";
        if (colIndex >= 0) {
            if (colIndex <= (ALPHABET_LENGTH - 1)) {
                resultColLabel = String.valueOf((char) (colIndex + 'A'));
            } else {
                resultColLabel = getColLabelFromColIndex(colIndex / 26 - 1) + getColLabelFromColIndex(colIndex % 26);
            }
        } else {
            throw new RuntimeException("列号必须大于0");
        }
        return resultColLabel;
    }
}
