import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;

import java.io.FileInputStream;
import java.io.IOException;

public class ReadExel {
    public static void main(String[] args) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream("exel.xls"));
        HSSFSheet sheet = workbook.getSheetAt(0);
        HSSFRow row = sheet.getRow(0);
        HSSFCell cell = row.getCell(0);
        HSSFCell cell1 = row.getCell(1);
        HSSFCell cell2 = row.getCell(2);
        HSSFCell cell3 = row.getCell(3);
        if (cell.getCellType() == CellType.STRING) {
            System.out.println(row.getCell(0).getStringCellValue());
        }
        if (cell1.getCellType()==CellType.NUMERIC){
            System.out.println(row.getCell(1).getDateCellValue());
        }
        if (cell2.getCellType()==CellType.STRING){
            System.out.println(row.getCell(2).getStringCellValue());
        }
    }
}
