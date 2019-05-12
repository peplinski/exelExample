import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class WriteExelMoreData {
    public static void main(String[] args) {

        HSSFWorkbook workbook = new HSSFWorkbook();

        HSSFSheet sheet = workbook.createSheet("Employee Data");
        sheet.autoSizeColumn(2);
        sheet.autoSizeColumn(3);
        sheet.autoSizeColumn(4);

        //Te dane muszą być zapisane
        Map<String, Object[]> data = new TreeMap<String, Object[]>();
        data.put("1", new Object[]{"NR", "Linia", "godz_rozpoczecia", "godz_zakonczenia"});
        data.put("2", new Object[]{ "1234", "K", "4:33", "6:45"});
        data.put("3", new Object[]{ "9999", "K", "05:23", "08:03"});
        data.put("4", new Object[]{ "8888", "R", "4:24", "13:00"});
        data.put("5", new Object[]{ "3333", "S", "05:46", "14:18"});

        //iteruje po data i zapisuję w arkuszu
        Set<String> keyset = data.keySet();
        int rownum = 0;
        for (String key : keyset) {
            Row row = sheet.createRow(rownum++);
            Object[] objects = data.get(key);
            int cellnum = 0;
            for (Object obj : objects) {
                Cell cell = row.createCell(cellnum++);
                if (obj instanceof String)
                    cell.setCellValue((String) obj);
                else if (obj instanceof Integer)
                    cell.setCellValue((Integer) obj);
            }
        }
        try {
            FileOutputStream out = new FileOutputStream(new File("wiecejDanych.xls"));
            try {
                workbook.write(out);
                out.close();
                System.out.println("Dane zostały zapisane na dysk");
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
    }
}
