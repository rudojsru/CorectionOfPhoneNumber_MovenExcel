package workWichExcel;

import workWichTextTxtFiles.TextWriterReader;

import java.io.IOException;
import java.util.List;
import java.util.Map;

public class AnalizatorExcelText {
    public Map  analizatorExcTex()throws IOException {
        ExcelReader excelReader = new ExcelReader();
        TextWriterReader textWriterReader = new TextWriterReader();
        List l = textWriterReader.listModifiPhone(); // List from TextWriter
        Map m = excelReader.exselWriterMap();  // Maps from Exsel Writer

        for (int i = 0; i < l.size(); i++) {     // List size for беребрать evry elements
            String listWordText = (String) l.get(i);
            //       System.out.println(listWordText+"-------------------");
            String numberOrTextFromList[] = listWordText.split("");
            if ((numberOrTextFromList[0].equals("0") != true) && (numberOrTextFromList[1].equals("7") != true)) {

                for (Object key : m.keySet()) { // взять новый обьект из  MAP
                    String mapWord = (String) key;
                    if (mapWord.equals(listWordText)) {  //in List exist FELDS name from Maps
                        listWordText = (String) l.get(++i);
                        numberOrTextFromList = listWordText.split("");
                        while ((numberOrTextFromList[0].equals("0") == true) && (numberOrTextFromList[1].equals("7") == true)) { // пока в списке телефон, то добавляем его в Мап


                            List l1 = (List<String>) m.get(mapWord);
                            l1.add(listWordText);
                            m.put(mapWord, l1);
                            if (i == l.size() - 1) {
                                break;
                            }

                            listWordText = (String) l.get(++i);
                            numberOrTextFromList = listWordText.split("");


                        }
                        i--;
                        break;
                    }

                }
            }
        }
     //   System.out.println(m);
        return m;
    }


}

