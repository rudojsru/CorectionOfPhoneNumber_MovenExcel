package workWichExcel;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import workWichTextTxt.TextWriterReader;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelWriter {
    public static void main(String[] args) throws IOException {
        TextWriterReader text =new TextWriterReader();
        String track = text.trackToFiles()+"ex3.xls";  // files trackl
        FileInputStream fis = new FileInputStream(track);
        Workbook wb = new HSSFWorkbook(fis);
        Sheet sheet = wb.getSheetAt(0);
//-----------------------
        int numer = sheet.getRow(0).getLastCellNum();
        Map<Integer, List<String>> data = new HashMap<>();
        System.out.println(sheet.getLastRowNum());
        for (int q = 0; q < numer; q++) {
            int cellIterator = 0;
            System.out.println("----------------");
            data.put(q, new ArrayList<String>());
            while (true) {
                data.get(q).add(getCellText(sheet.getRow(cellIterator).getCell(q)));
                cellIterator++;
                if (sheet.getRow(cellIterator) == null || sheet.getRow(cellIterator).getCell(q) == null) {
                    System.out.println("STOP");
                    break;
                }
            }
            System.out.println(data.get(q));
        }

        fis.close();



    }


    public static String getCellText(Cell cell) {
        // Alternatively, get the value and format it yourselfcellIteratorcellIteratorcellIteratorcellIterator

        String result = "";
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                result = cell.getRichStringCellValue().getString();
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    result = cell.getDateCellValue().toString();
                } else {
                    result = String.valueOf(cell.getNumericCellValue());
                }
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                result = String.valueOf(cell.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_FORMULA:
                result = cell.getCellFormula();
                break;

            default:
                break;
        }
        return result;
    }
}
