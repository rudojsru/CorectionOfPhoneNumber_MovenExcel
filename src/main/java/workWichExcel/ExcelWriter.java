package workWichExcel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import workWichTextTxtFiles.TextWriterReader;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

public class ExcelWriter {
    public static void main(String[] args) throws IOException {
        AnalizatorExcelText mapNew = new AnalizatorExcelText();
        Map m = mapNew.analizatorExcTex();  // наш сложеный из екселя и текста массив
        System.out.println(m);
        TextWriterReader text = new TextWriterReader();
        String track = text.trackToFiles() + "ex6.xls";  // files trackl

        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("processedFiles");
        Row row;

        int i = 0;
        for (Object list : m.keySet()) {
             List list1 = (List) m.get(list);
            for (int k = 0; k < list1.size(); k++) {
                if(sheet.getRow(k) == null ){

                    String textFromList = (String) list1.get(k);
                    Cell cell = sheet.createRow(k).createCell(i,CellType.STRING);
                    cell.setCellValue(textFromList);
                }else {
                    String textFromList = (String) list1.get(k);
                    Cell cell = sheet.getRow(k).createCell(i,CellType.STRING);
                    cell.setCellValue(textFromList);
                }


            } i++;
         }
        FileOutputStream fileOut= new FileOutputStream(track);
        wb.write(fileOut);
        fileOut.close();
        wb.close();
    }
}

