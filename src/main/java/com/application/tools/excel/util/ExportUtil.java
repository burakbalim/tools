package com.application.tools.excel.util;

import com.application.tools.excel.annotation.ExcelAnnotation;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.List;
import java.util.stream.Collectors;

public class ExportUtil {

    public static <T> List<String> orderedHeader(Class<T> clazz) {
        List<String> headers = new ArrayList<>();
        Field[] fields = clazz.getDeclaredFields();
        List<Field> annotatedFields = Arrays.stream(fields)
                .filter(field -> field.isAnnotationPresent(ExcelAnnotation.class))
                .sorted(Comparator.comparingInt(field -> field.getAnnotation(ExcelAnnotation.class).order()))
                .toList();
        annotatedFields.forEach(field -> headers.add(field.getAnnotation(ExcelAnnotation.class).header()));
        return headers;
    }

    public static <T> List<Field> orderedFields(Class<T> clazz) {
        return Arrays.stream(clazz.getDeclaredFields())
                .filter(field -> field.isAnnotationPresent(ExcelAnnotation.class))
                .sorted(Comparator.comparingInt(field -> field.getAnnotation(ExcelAnnotation.class).order()))
                .collect(Collectors.toList());
    }
}
