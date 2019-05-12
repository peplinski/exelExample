import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;

import java.io.*;
import java.util.Iterator;

public class ReadExelMoreData {
    public static void main(String[] args) {

        try {
            FileInputStream file = new FileInputStream(new File("wiecejDanych.xls"));

            HSSFWorkbook workbook = new HSSFWorkbook(file);
//            System.out.println(workbook.getNumberOfSheets());
//            System.out.println(workbook.getSheet("export"));
//            System.out.println(workbook.getActiveSheetIndex());
//            System.out.println(workbook.getAllNames());


            HSSFSheet sheet = workbook.getSheetAt(0);



            //Iteruj po każdym rzędzie jeden po drugim
            Iterator<Row> rowIterator = sheet.iterator();

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                // Dla każdego wiersza przeglądaj wszystkie kolumny
                Iterator<Cell> cellIterator = row.cellIterator();

                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    String text = "";
                    //Sprawdź odpowiednio typ komórki i format
                    switch (cell.getCellType()) {

                        case NUMERIC:
                            if (DateUtil.isCellDateFormatted(cell)) {
                                text = String.valueOf(cell.getDateCellValue());
                            } else {
                                text = String.valueOf(cell.getNumericCellValue());
                            }
                            System.out.println(text);
                            break;
                        case STRING:
                            System.out.println(cell.getStringCellValue());
                            break;
                    }

                }
                System.out.println("");
            }
            file.close();

        }  catch (IOException e) {
            e.printStackTrace();
        }
    }
}
