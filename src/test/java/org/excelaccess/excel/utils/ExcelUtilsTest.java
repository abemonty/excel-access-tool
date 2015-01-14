package org.excelaccess.excel.utils;

import static org.junit.Assert.assertEquals;

import org.excelaccess.excel.utils.ExcelUtils;
import org.junit.Test;

public class ExcelUtilsTest {

    @Test
    public void computeColumnLetter() {
        assertEquals(0, ExcelUtils.computeColumnIndexFromLetters("A"));
        assertEquals(1, ExcelUtils.computeColumnIndexFromLetters("B"));
        assertEquals(7, ExcelUtils.computeColumnIndexFromLetters("H"));
        assertEquals(25, ExcelUtils.computeColumnIndexFromLetters("Z"));
        assertEquals(26, ExcelUtils.computeColumnIndexFromLetters("AA"));
        assertEquals(26 + 26 - 1, ExcelUtils.computeColumnIndexFromLetters("AZ"));
        // 26 * 2 + 24
        assertEquals(26 * 2 + 24 - 1, ExcelUtils.computeColumnIndexFromLetters("BX"));
    }
}
