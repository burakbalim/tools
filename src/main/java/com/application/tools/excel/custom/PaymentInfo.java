package com.application.tools.excel.custom;

import com.application.tools.excel.annotation.ExcelAnnotation;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;

import lombok.Setter;

@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
public class PaymentInfo {

    @ExcelAnnotation(header = "user_id", order = 1)
    public Integer userId;

    @ExcelAnnotation(header = "user_type", order = 2)
    public String userType;

    @ExcelAnnotation(header = "total_amount", order = 3)
    public Double totalAmount;

    @ExcelAnnotation(header = "currency", order = 5)
    public String currency;
}
