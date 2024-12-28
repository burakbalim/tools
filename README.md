# tools
Custom tools to be used on the projects

# PoiExcelWriterService

This project demonstrates how to use Apache POI to generate Excel files dynamically using annotations. The `PoiExcelWriterServiceImpl` class is an implementation of the `ExcelWriterService` interface that enables the creation of Excel files based on a list of objects and a defined annotation.

## Features

- Dynamically generates Excel headers and rows based on the annotated fields in the class.
- Supports multiple data types such as `String`, `Integer`, `Long`, `Float`, and `Double`.
- Uses Apache POI for creating `.xls` files.
- Customizable sheet names and column orders through annotations.

## How It Works

### 1. Define the Annotation

Use the `@ExcelAnnotation` to specify the header name and order for the columns in the Excel file.

```java
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelAnnotation {

    String header();
    int order();
}
```

### 2. Service Implementation

The core logic for writing data to an Excel file is in the `PoiExcelWriterServiceImpl` class. This class includes the following methods:

- **`writeContent`**: Creates the Excel sheet, headers, and rows from the provided data list.
- **`createRowHeader`**: Populates the header row using the annotation-defined headers.
- **`createRow`**: Populates data rows based on the list of objects.
- **`retrieveContent`**: Converts the workbook into a byte array for easy export or saving.

### 3. Example Usage

Here is an example of how to use the service to generate an Excel file:

#### Define the Model Class

```java
public class PaymentInfo {
    @ExcelAnnotation(header = "ID", order = 1)
    private int id;

    @ExcelAnnotation(header = "Customer Name", order = 2)
    private String customerName;

    @ExcelAnnotation(header = "Amount", order = 3)
    private Double amount;

    @ExcelAnnotation(header = "Currency", order = 4)
    private String currency;

    public PaymentInfo(int id, String customerName, Double amount, String currency) {
        this.id = id;
        this.customerName = customerName;
        this.amount = amount;
        this.currency = currency;
    }
}
```

#### Generate the Excel File

```java
public void test() throws IOException {
    List<PaymentInfo> paymentInfos = new ArrayList<>();

    for (int i = 1; i <= 5; i++) {
        paymentInfos.add(new PaymentInfo(i, "Customer" + i, 30.0 * i, "USD"));
    }

    Workbook workbook = new HSSFWorkbook(); // Using HSSFWorkbook for .xls
    byte[] excelData = excelWriterService.writeContent(workbook, new ExcelWriterModel<>(paymentInfos, PaymentInfo.class, "Test"));

    try (FileOutputStream fileOutputStream = new FileOutputStream("test.xls")) {
        fileOutputStream.write(excelData);
    }
}
```

## Prerequisites

- **Java 8 or later**
- Apache POI Library

Add the following Maven dependency for Apache POI:

```xml
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi</artifactId>
    <version>5.2.3</version>
</dependency>
```

## How to Run

1. Clone this repository.
2. Add your model class with annotated fields using `@ExcelAnnotation`.
3. Use the `PoiExcelWriterServiceImpl` service to generate your Excel file.
4. Run the `test` method to see the generated file in your project directory.

## Notes

- The order of the columns in the Excel file is determined by the `order` parameter in the `@ExcelAnnotation`.
- Ensure that all required fields in the model class are annotated correctly to avoid missing data in the Excel file.

## License

This project is open-source and available under the MIT License.

