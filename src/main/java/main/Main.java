package main;

import interfaces.CanDo;
import interfaces.ExcelReader;

public class Main {

    private static final String fileName = "test1.xls";
    private static final String cellText = "Евро";
    private static final String imageName = "default.png";
    private static final String sheetName = "List22";
    private static final String cellName = "A1";
    private static final boolean equalsOrContains = false;


    public static void main(String[] args) {
        CanDo excelReader = new ExcelReader(fileName);

        excelReader.findText(cellText, equalsOrContains);
        excelReader.findText("Дзю", false);
        // excelReader.insertImage(fileName, imageName, sheetName, cellName);

    }
}
