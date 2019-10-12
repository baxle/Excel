package Interface;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.*;

import org.apache.log4j.Logger;
import org.apache.poi.ss.util.WorkbookUtil;

public class ExcelReader implements CanDo {

    private Workbook workbook;
    private int textCount;
    private final Logger logger = Logger.getLogger(ExcelReader.class);
    private int listCount;


    private void openFile(String fileName) {
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
    }

    /**
     * Функция поиска текста в эксель файле.
     *
     * @param fileName         - имя эксель файла.
     * @param text             - текст, который ищем в файле.
     * @param equalsOrContains - true - полное совпадение @param text c текстом в ячейке, false - частичное.
     */
    @Override
    public void findText(String fileName, String text, boolean equalsOrContains) {
        textCount = 0;
        openFile(fileName);

        workbook.forEach(sheet -> {
            sheet.forEach(row -> {
                row.forEach(cell -> {
                    if (cell.getCellTypeEnum() == CellType.STRING) {
                        if (cell.getStringCellValue().contains(text) && !equalsOrContains) {
                            System.out.printf("Текст %s найден в листе %s в ячейке %s.\n", text, cell.getSheet().getSheetName(), cell.getAddress());
                            textCount++;
                        }
                        if (cell.getStringCellValue().equals(text) && equalsOrContains) {
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
        try {
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Override
    public void insertImage(String fileName, String imageName, String sheetName, String cell) {

        openFile(fileName);

        workbook.forEach(sheet -> {
            if (sheet.getSheetName().equals(sheetName)) {
                listCount++;
            }
        });


        System.out.println(listCount);
        if (listCount == 0) {
            System.err.printf("Листа %s не найдено. Создаем новый лист\n", sheetName);


            String safeName = WorkbookUtil.createSafeSheetName(sheetName); // returns " O'Brien's sales   "
            Sheet sheet3 = workbook.createSheet(safeName);
            try (OutputStream fileOut = new FileOutputStream(fileName)) {
                workbook.write(fileOut);
            } catch (IOException e) {
                e.printStackTrace();
            }
        }


        try {
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


}
