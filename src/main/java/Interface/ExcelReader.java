package Interface;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.log4j.Logger;

public class ExcelReader implements CanDo {

    private Workbook workbook;
    private int textCount;
    private final Logger logger = Logger.getLogger(ExcelReader.class);


    /**
     * Функция поиска текста в эксель файле.
     * @param fileName - имя эксель файла.
     * @param text - текст, который ищем в файле.
     * @param equalsOrContains - true - полное совпадение @param text c текстом в ячейке, false - частичное.
     */
    @Override
    public void findText(String fileName, String text, boolean equalsOrContains) {

        try {
            workbook = WorkbookFactory.create(new File(fileName));
        } catch (FileNotFoundException e) {
            System.err.printf("Файла %s не существует.\n", fileName);
            logger.error("Заданного файла " + fileName + " не существует.");
            System.exit(0);
        } catch (IOException e) {
            e.printStackTrace();
            logger.error("Ошибка.");
        } catch (InvalidFormatException e) {
            e.printStackTrace();
            logger.error("Ошибка InvalidFormatException.");
        }

        workbook.forEach(sheet -> {
            sheet.forEach(row -> {
                row.forEach(cell -> {
                    if (cell.getCellTypeEnum() == CellType.STRING) {
                            if (cell.getStringCellValue().contains(text)&&!equalsOrContains) {
                                System.out.printf("Текст %s найден в листе %s в ячейке %s.\n", text, cell.getSheet().getSheetName(), cell.getAddress());
                                textCount++;
                        }
                        if (cell.getStringCellValue().equals(text)&&equalsOrContains) {
                            System.out.printf("Текст %s найден в листе %s в ячейке %s.\n", text, cell.getSheet().getSheetName(), cell.getAddress());
                            textCount++;
                        }
                    }
                });
            });
        });
        if (textCount == 0) {
            System.err.printf("Текст %s не найден.\n", text);
            logger.error("Искомого текста " + text + " в файле не найдено.");
        }
    }
}
