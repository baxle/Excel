package Main;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.log4j.Logger;

public class ExcelReader {

    private static Workbook workbook;
    private static final String EXCEL_FILE_PATH = "test1.xls";
    private static String cellText = "Евро";
    private static int textCount;
    final static Logger logger = Logger.getLogger(ExcelReader.class);



    public static void main(String[] args)  {

        try {
            workbook = WorkbookFactory.create(new File(EXCEL_FILE_PATH));
        } catch (FileNotFoundException e){
            System.err.printf("Файла %s не существует.", EXCEL_FILE_PATH);
            logger.error("Заданного в EXCEL_FILE_PATH файла не существует.");
            System.exit(0);
        }
        catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }

        /**
         * Поиск текста {@link #cellText} по всем ячейкам файла {@link #EXCEL_FILE_PATH}
         */
        workbook.forEach(sheet -> {
            sheet.forEach(row -> {
                row.forEach(cell -> {
                    if(cell.getCellTypeEnum() == CellType.STRING) {
                        if (cell.getStringCellValue().contains(cellText)) {
                            // для полного совпадения //if (cell.getStringCellValue().equals(cellText)) {
                            System.out.printf("Текст %s найден в листе %s в ячейке %s.\n", cellText, cell.getSheet().getSheetName(), cell.getAddress());
                            textCount++;
                        }
                    }
                });
            });
        });
        if (textCount == 0) {
            System.err.printf("Текст %s не найден.", cellText);
            logger.error("Это сообщение ошибки");
        }


    }
}
