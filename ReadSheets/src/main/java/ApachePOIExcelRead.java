/**********************************************************************.
 *                                                                     *
 *         Copyright (c) Ultra Electronics Airport Systems 2018     *
 *                         All rights reserved                         *
 *                                                                     *
 ***********************************************************************/


import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class ApachePOIExcelRead {
  public static final String SAMPLE_XLSX_FILE_PATH = "C:\\source\\ExcelReader\\ReadSheets\\src\\test\\resources\\JavaInterfaceMapping.xls";

  public static void main(String[] args) throws IOException, InvalidFormatException {

    // Creating a Workbook from an Excel file (.xls or .xlsx)
    Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));

    // Retrieving the number of sheets in the Workbook
    System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");

        /*
           =============================================================
           Iterating over all the sheets in the workbook (Multiple ways)
           =============================================================
        */

   /* // 1. You can obtain a sheetIterator and iterate over it
    Iterator<Sheet> sheetIterator = workbook.sheetIterator();
    System.out.println("Retrieving Sheets using Iterator");
    while (sheetIterator.hasNext()) {
      Sheet sheet = sheetIterator.next();
      System.out.println("=> " + sheet.getSheetName());
    }

    // 2. Or you can use a for-each loop
    System.out.println("Retrieving Sheets using for-each loop");
    for(Sheet sheet: workbook) {
      System.out.println("=> " + sheet.getSheetName());
    }*/

    // 3. Or you can use a Java 8 forEach with lambda
    System.out.println("Retrieving Sheets using Java 8 forEach with lambda");
    workbook.forEach(sheet -> {
      System.out.println("=> " + sheet.getSheetName());
    });

    //Sheet sheet = workbook.getSheetAt(0);
    DataFormatter dataFormatter = new DataFormatter();
    // 3. Or you can use Java 8 forEach loop with lambda
    System.out.println("\n\nIterating over Rows and Columns using Java 8 forEach with lambda\n");
   ArrayList<String> props = new ArrayList<>();
    for(Sheet sht:workbook){
      for(Row r:sht) {
        String rowVal= "";
        for(Cell c:r){
          String cellValue = dataFormatter.formatCellValue(c);
          if (c.getColumnIndex() == 0) {
            rowVal=cellValue;
          }else if(c.getColumnIndex() == 1){
            rowVal+=".CURRENT_"+sht.getSheetName().toUpperCase()+"."+cellValue+"=";
          }else{
            if(cellValue.contains("Map to ")||cellValue.contains("Mapped to ")){
              rowVal+=cellValue.substring(cellValue.length()-3,cellValue.length());
            }else if(cellValue.contains("Derive")){
              rowVal = "#"+rowVal+cellValue;
            }else{
              rowVal+=cellValue;
            }
          }
        }
        props.add(rowVal);
      }
    }
    props.forEach((t)-> System.out.println(t));
    /*workbook.forEach(sheet -> {
          sheet.forEach(row -> {

            row.forEach(cell -> {
              String cellValue = dataFormatter.formatCellValue(cell);
              if (cell.getColumnIndex() == 0) {
                System.out.print("CURRENT_"+sheet.getSheetName().toUpperCase()+"."+cellValue+"=");
                rowVal="CURRENT_"+sheet.getSheetName().toUpperCase()+"."+cellValue+"=";
              }else{
                if(cellValue.contains("Map to ")){
                  System.out.print(cellValue.replace("Map to ",""));
                  rowVal+=cellValue.replace("Map to ","");
                }else{
                  System.out.print(cellValue);
                }
              }

            });
            System.out.println();
          });
        });*/
    // Closing the workbook
    workbook.close();
        /*
           ==================================================================
           Iterating over all the rows and columns in a Sheet (Multiple ways)
           ==================================================================
        */

    // Getting the Sheet at index zero

/*
    // Create a DataFormatter to format and get each cell's value as String
    DataFormatter dataFormatter = new DataFormatter();

    // 1. You can obtain a rowIterator and columnIterator and iterate over them
    System.out.println("\n\nIterating over Rows and Columns using Iterator\n");
    Iterator<Row> rowIterator = sheet.rowIterator();
    while (rowIterator.hasNext()) {
      Row row = rowIterator.next();

      // Now let's iterate over the columns of the current row
      Iterator<Cell> cellIterator = row.cellIterator();

      while (cellIterator.hasNext()) {
        Cell cell = cellIterator.next();
        String cellValue = dataFormatter.formatCellValue(cell);
        System.out.print(cellValue + "\t");
      }
      System.out.println();
    }

    // 2. Or you can use a for-each loop to iterate over the rows and columns
    System.out.println("\n\nIterating over Rows and Columns using for-each loop\n");
    for (Row row: sheet) {
      for(Cell cell: row) {
        String cellValue = dataFormatter.formatCellValue(cell);
        System.out.print(cellValue + "\t");
      }
      System.out.println();
    }*/


  }

  /*private static void printCellValue(Cell cell) {
    switch (cell.getCellTypeEnum()) {
      case BOOLEAN:
        System.out.print(cell.getBooleanCellValue());
        break;
      case STRING:
        System.out.print(cell.getRichStringCellValue().getString());
        break;
      case NUMERIC:
        if (DateUtil.isCellDateFormatted(cell)) {
          System.out.print(cell.getDateCellValue());
        } else {
          System.out.print(cell.getNumericCellValue());
        }
        break;
      case FORMULA:
        System.out.print(cell.getCellFormula());
        break;
      case BLANK:
        System.out.print("");
        break;
      default:
        System.out.print("");
    }

    System.out.print("\t");
  }*/
}