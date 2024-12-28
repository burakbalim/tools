package com.application.tools;

import com.application.tools.excel.*;
import com.application.tools.excel.custom.PaymentInfo;
import com.application.tools.excel.model.ExcelWriterModel;
import com.application.tools.excel.util.WorkbookFactory;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

@SpringBootApplication
@RequiredArgsConstructor
public class ToolsApplication implements CommandLineRunner {

	private final ExcelWriterService excelWriterService;

	public static void main(String[] args) {
		SpringApplication.run(ToolsApplication.class, args);
	}

	@Override
	public void run(String... args) throws Exception {
		List<PaymentInfo> paymentInfos = new ArrayList<>();

		for (int i= 1; i <= 5; i++) {
			paymentInfos.add(new PaymentInfo( i, "Customer", 30.0, "USD"));
		}

		Workbook workbook = WorkbookFactory.createWorkbook("xls");
		byte[] tests = excelWriterService.writeContent(workbook, new ExcelWriterModel<>(paymentInfos, PaymentInfo.class, "Test"));

		FileOutputStream fileOutputStream = new FileOutputStream("test.xls");

		fileOutputStream.write(tests);

		fileOutputStream.close();
	}
}
