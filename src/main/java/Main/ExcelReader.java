package Main;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import java.io.File;
import java.io.IOException;
import java.util.List;

public class ExcelReader {

    public static final String EXCEL_FILE_PATH = "test.xls";
    private static String cellText = "лист3";
    private static int textCount;



    public static void main(String[] args) throws IOException, InvalidFormatException {

        Workbook workbook = WorkbookFactory.create(new File(EXCEL_FILE_PATH));

        /**
         * Поиск текста {@link #cellText} по всем ячейкам файла {@link #EXCEL_FILE_PATH}
         */
        workbook.forEach(sheet -> {
            sheet.forEach(row -> {
                row.forEach(cell -> {
                    if (cell.getStringCellValue().equals(cellText)) {
                        System.out.printf("Текст %s найден в листе %s в ячейке %s.\n", cellText, cell.getSheet().getSheetName(), cell.getAddress());
                        textCount++;
                    }
                });
            });
        });
        if (textCount == 0) {
            System.err.printf("Текст %s не найден.", cellText);
        }


    }
}
