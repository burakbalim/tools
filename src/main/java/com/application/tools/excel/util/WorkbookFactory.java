package com.application.tools.excel.util;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

public class WorkbookFactory {

    public static Workbook createWorkbook(String type) {
        switch (type.toLowerCase()) {
            case "xls":
                return new HSSFWorkbook();
            default:
                throw new IllegalArgumentException("Unsupported workbook type: " + type);
        }
    }
}
