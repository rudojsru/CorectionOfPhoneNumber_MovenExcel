package workWichExcel;

import workWichTextTxtFiles.TextWriterReader;

import java.io.IOException;
import java.util.List;
import java.util.Map;

public class AnalizatorExcelText {


    public static void main(String[] args) throws IOException {
        ExcelWriter excelWriter = new ExcelWriter();
        TextWriterReader textWriterReader = new TextWriterReader();
        List l = textWriterReader.listModifiPhone();
        Map m =excelWriter.exselWriterMap();

        String mapWord;
        for (int i = 0; i < l.size(); i++) {
            String listWord=(String) l.get(i);
            System.out.println(listWord+"-------------------");
            String numberOrTextFromList[] = listWord.split("");

            if ((numberOrTextFromList[0].equals("0") != true) && (numberOrTextFromList[1].equals("7") != true)) {

                for (Object key: m.keySet()) {
                    mapWord= (String) key;
                    if (mapWord.equals(listWord)){
                        System.out.println(mapWord+"=="+listWord);
                    }
                }

            }
        }

    }
}
