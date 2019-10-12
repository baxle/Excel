package Main;

import Interface.CanDo;
import Interface.ExcelReader;

public class Main {

    private static final String fileName = "test1.xls";
    private static final String cellText = "Евро";
    private static final String imageName = "default.png";
    private static final String sheetName = "List22";
    private static final String cellName = "A1";
    private static final boolean equalsOrContains = false;


    public static void main(String[] args) {
        CanDo excelReader = new ExcelReader();

        excelReader.findText(fileName, cellText, equalsOrContains);
        excelReader.findText(fileName, "Дзю", false);
        excelReader.insertImage(fileName, imageName, sheetName, cellName);
    }
}
