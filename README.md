# simple-spreadsheet

This is a small library to create very simple spreadsheets in .xlsx format.  No 3rd-party libraries are required.  It's
useful where you'd normally just write a .csv file for loading in to Excel but you want to control the formatting of numbers.  Also, numbers entered as text have the "number entered as text" warning message suppressed and leading zeros are preserved.

Three data-types are supported:

* String
* BigDecimal
* java.util.Date

The formatting of the BigDecimal is according to the number of decimals which is obtained from calling scale().  The Date format is YYYY-MM-DD.  String cells are written as given.

Multiple worksheets are supported.

## Simple Usage

```java
        Workbook wb = new Workbook();
        Worksheet ws = wb.addWorksheet("My First Sheet");
        
        ws.nextRow()
                .append("Hello")
                .append(new BigDecimal("100.00"))
                .append(new Date())
                .append("001");
        ws.newRow(3) // specify a row number to override the next number
                .append("World")
                .append(new BigDecimal("0.0"))
                .append(new Date());
        
        ws = wb.addWorksheet("My Other Sheet");
        ws.nextRow()
                .append("Hello World")
                .skipColumn() // adds a blank column
                .append(new BigDecimal("9.900"))
                .append(new Date());
        
        try {
            wb.write("test.xlsx");
        } catch (IOException ex) {
            Logger.getLogger(WorkbookTest.class.getName()).log(Level.SEVERE, null, ex);
        }
```
