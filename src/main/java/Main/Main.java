package Main;

import Interface.CanDo;
import Interface.ExcelReader;

public class Main {

    private static final String fileName = "test1.xls";
    private static String cellText = "Евро";
    private static boolean equalsOrContains = false;


    public static void main(String[] args) {
        CanDo excelReader = new ExcelReader();

        excelReader.findText(fileName, cellText, equalsOrContains);
    }
}
