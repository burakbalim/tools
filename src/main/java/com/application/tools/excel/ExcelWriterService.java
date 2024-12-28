package com.application.tools.excel;

import com.application.tools.excel.model.ExcelWriterModel;
import org.apache.poi.ss.usermodel.Workbook;

public interface ExcelWriterService {

    <T> byte[] writeContent(Workbook workbook, ExcelWriterModel<T> model);
}
