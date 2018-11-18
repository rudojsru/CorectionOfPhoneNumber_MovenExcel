package workWichExcel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import workWichTextTxtFiles.TextWriterReader;

import java.io.*;
import java.util.List;

public class Main {
    public static void main(String[] args) throws IOException {

        TextWriterReader text = new TextWriterReader();
        File track = new File(text.trackToFiles() + "ex3.xls");
        File track2 = new File(text.trackToFiles() + "modifiedExcel.xls");
        Main.copyFileUsingStream(track, track2); // - creirt copi of files


        FileInputStream fis = new FileInputStream(track2);
        Workbook wb = new HSSFWorkbook(fis);
        Sheet sheet = wb.getSheetAt(0);
        int numerCell = sheet.getRow(0).getLastCellNum();
        Row row = sheet.getRow(0);

        TextWriterReader textWriterReader = new TextWriterReader();
        List l = textWriterReader.listModifiPhone(); // List from TextWriter
        System.out.println(l);
        for (int i = 0; i < l.size(); i++) {     // List size for беребрать evry elements
            String listWordText = l.get(i).toString().toLowerCase();

            String numberOrTextFromList[] = listWordText.split("");
            if ((numberOrTextFromList[0].equals("0") != true) && (numberOrTextFromList[1].equals("7") != true)) {
                for (int c = 0; c < numerCell; c++) {   // потом в С поставить 1   с=1
                    String nameList = String.valueOf(sheet.getRow(0).getCell(c)).toLowerCase();
                    if (listWordText.equals(nameList)) {
                        int r = 0;

                        while ((sheet.getRow(r) != null) && (sheet.getRow(r).getCell(c) != null)
                                && (!"".equals(String.valueOf(sheet.getRow(r).getCell(c))))) {   // находит конец столбца в таблице ексель

                            r++;

                        }

                        i++;
                        listWordText = (String) l.get(i);
                        String numberOrTextFromList2[] = listWordText.split("");

                        //для изменения цвета добовляемых ячеек
                        CellStyle style = wb.createCellStyle();
                        Font font = wb.createFont();
                        font.setColor(HSSFColor.GREEN.index);
                        style.setFont(font);
                        // Добавляют в ексель телефоны из файла текст.

                        while ((numberOrTextFromList2[0].equals("0") == true) && (numberOrTextFromList2[1].equals("7") == true)
                                && (!"null".equals(listWordText))) {

                            Cell cell = sheet.getRow(r).createCell(c);
                            cell.setCellStyle(style);
                            cell.setCellValue(listWordText);
                            r++;
                            i++;
                            if (i < l.size()) {
                                listWordText = (String) l.get(i);
                                numberOrTextFromList2 = listWordText.split("");
                            } else break;

                        }

                        i--;
                        break;
                    }
                }
            }
        }
        fis.close();
        FileOutputStream output_file = new FileOutputStream(track2);
        //write changes
        wb.write(output_file);
        output_file.close();
        System.out.println("The End");


        SearchDoublePhone s = new SearchDoublePhone();
        s.searchDouble();
    }


    private static void copyFileUsingStream(File source, File dest) throws IOException {


        InputStream is = null;
        OutputStream os = null;
        try {
            is = new FileInputStream(source);
            os = new FileOutputStream(dest);
            byte[] buffer = new byte[1024];
            int length;
            while ((length = is.read(buffer)) > 0) {
                os.write(buffer, 0, length);
            }
        } finally {
            is.close();
            os.close();
        }
    }
}
