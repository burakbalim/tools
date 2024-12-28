package com.application.tools.excel.impl;

import com.application.tools.excel.ExcelWriterService;
import com.application.tools.excel.model.ExcelWriterModel;
import com.application.tools.excel.util.ExportUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.List;
import java.util.Objects;

@Slf4j
@Service
public class PoiExcelWriterServiceImpl implements ExcelWriterService {

    public <T> byte[] writeContent(Workbook workbook, ExcelWriterModel<T> model)  {
        Sheet sheet = workbook.createSheet(model.getSheetName());
        createRowHeader(sheet, ExportUtil.orderedHeader(model.getClazz()));
        createRow(sheet, model.getList(), model.getClazz());
        return retrieveContent(workbook);
    }

    private byte[] retrieveContent(Workbook workbook) {
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        try {
            workbook.write(byteArrayOutputStream);
            byteArrayOutputStream.close();
        } catch (IOException e) {
            log.error("Error occurred while getting bytes of xls file.", e);
            throw new RuntimeException("Error reading bytes from xls file");
        }
        return byteArrayOutputStream.toByteArray();
    }

    private <T> void createRow(Sheet sheet, List<T> datalist, Class<T> clazz) {
        int rowNum = 1;
        List<Field> fields = ExportUtil.orderedFields(clazz);
        for (T item : datalist) {
            Row row = sheet.createRow(rowNum++);
            int colNum = 0;
            assert fields != null;
            for (Field field : fields) {
                Cell cell = row.createCell(colNum++);
                setCellValue(field, item, cell);
            }
        }
    }

    private void createRowHeader(Sheet sheet, List<String> datalist) {
        Row row = sheet.createRow(0);
        sheet.autoSizeColumn(0);
        int colNum = 0;
        for (String value : datalist) {
            Cell cell = row.createCell(colNum++);
            cell.setCellValue(value);
        }
    }

    private void setCellValue(Field field, Object instance, Cell cell)  {
        Class<?> type = field.getType();
        Object value;
        try {
            value = field.get(instance);
        } catch (IllegalAccessException e) {
            throw new RuntimeException("Occurred exception while setting cell values", e);
        }
        if (Objects.isNull(value)) {
        }
        else if (type.equals(String.class)) {
            cell.setCellValue((String) value);
        } else if (type.equals(Integer.class)) {
            cell.setCellValue((Integer) value);
        } else if (type.equals(Long.class)) {
            cell.setCellValue((Long) value);
        } else if (type.equals(Float.class)) {
            cell.setCellValue((Float) value);
        } else if (type.equals(Double.class)) {
            cell.setCellValue((Double) value);
        } else if (type.equals(Short.class)) {
            cell.setCellValue((short) value);
        }
    }
}
