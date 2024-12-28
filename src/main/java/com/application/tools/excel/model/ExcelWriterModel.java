package com.application.tools.excel.model;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;

import java.util.List;

@Getter
@Setter
@AllArgsConstructor
public class ExcelWriterModel<T> {

    private List<T> list;
    private Class<T> clazz;
    private String sheetName;
}
