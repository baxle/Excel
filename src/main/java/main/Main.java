package main;

import interfaces.CanDo;
import interfaces.ExcelReader;

public class Main {

    private static final String fileName = "test1.xls";
    private static final String cellText = "Евро2525";
    private static final String imageName = "test.jpg";
    private static final String sheetName = "List144";
    private static final String cellName = "A1";
    private static final boolean equalsOrContains = true;


    public static void main(String[] args) {
        CanDo excelReader = new ExcelReader(fileName);
        ExcelReader excelReader1 = new ExcelReader(fileName);

        excelReader.findText(cellText, equalsOrContains);
        excelReader.findText("Дзю", false);
        // excelReader.insertImage(fileName, imageName, sheetName, cellName);

     excelReader1.addImage(sheetName, imageName);




    }
}
