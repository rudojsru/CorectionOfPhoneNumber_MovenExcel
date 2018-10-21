package workWichExcel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
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

        int j = 0;  // cell
        for (Object list : m.keySet()) {
            int i = 0;   //row
            List list1 = (List) m.get(list);
            for (int k = 0; k < list1.size(); k++) {
                // System.out.println(list1);
                String textFromList = (String) list1.get(k);
                System.out.print(textFromList);
                Cell cell = sheet.createRow(i).createCell(j);
                cell.setCellValue(textFromList);
            //    wb.write(new FileOutputStream(track));
                i++;

            } j++;
            System.out.println();
        }
        FileOutputStream fileOut= new FileOutputStream(track);
        wb.write(fileOut);
        fileOut.close();
      //  wb.write(new FileOutputStream(track));
        wb.close();
    }
}

