package workWichExcel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import workWichTextTxtFiles.TextWriterReader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;

public class SearchDoublePhone {

    public static void main(String[] args) throws IOException {
        searchDouble();
    }

    public static void searchDouble() throws IOException {
        TextWriterReader text = new TextWriterReader();

        File track2 = new File(text.trackToFiles() + "modifiedExcel.xls");

        FileInputStream fis = new FileInputStream(track2);
        Workbook wb = new HSSFWorkbook(fis);
        Sheet sheet = wb.getSheetAt(0);
        int numerCell = sheet.getRow(0).getLastCellNum();


        for (int c = 1; c < numerCell; c++) {
            int r = 0;
            List list = new ArrayList<>();
            while (sheet.getRow(r) != null && sheet.getRow(r).getCell(c) != null
                    && !"".equals(String.valueOf(sheet.getRow(r).getCell(c)))) {

                String nameCell = String.valueOf(sheet.getRow(r).getCell(c));
                list.add(nameCell);
                r++;
            }

            ArrayList dublicat= arrayCheck(list);

            if(dublicat.size()>0){
                CellStyle style = wb.createCellStyle();
                Font font = wb.createFont();
                font.setColor(HSSFColor.RED.index);
                style.setFont(font);
                for (Object d: dublicat ) {
                     Cell cell = sheet.getRow((int)d).getCell(c);
                    cell.setCellStyle(style);
                }

            }

        }

        fis.close();
        FileOutputStream output_file = new FileOutputStream(track2);
        //write changes
        wb.write(output_file);
        output_file.close();
        System.out.println("The End 2");
    }
    public static ArrayList  arrayCheck(List list){ // метод находит повторение телефона и указывает в какой позиции он находиться.
        HashSet<String> used = new HashSet<>();
        ArrayList<Integer> dublicat = new ArrayList<>();

        for(int i = 0; i < list.size(); i++){
            if(used.contains(list.get(i))){
                continue;
            } else
                used.add((String) list.get(i));

            ArrayList<Integer> positions = new ArrayList<>();
            positions.add(i);
            int dublicatNumber=0;
            for(int j = i + 1; j < list.size(); j++){

                if(list.get(i) == list.get(j)){
                    positions.add(j);
                    dublicatNumber++;
                }
            }

            if(dublicatNumber>0){
                System.out.println(list.get(i) +" wstreczaetsia w posicii");
                for (Integer p: positions){
                    System.out.println(p+ ", ");
                    dublicat.add(p);
                }
            }

        }
        return dublicat;
    }
}
