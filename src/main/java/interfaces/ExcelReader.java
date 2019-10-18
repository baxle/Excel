package interfaces;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.net.URL;
import java.util.function.BiPredicate;

import org.apache.log4j.Logger;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.util.IOUtils;


public class ExcelReader implements CanDo {

    private Workbook workbook;
    private int textCount;
    private final Logger logger = Logger.getLogger(ExcelReader.class);
    private int listCount;
    private String fileName;


    public ExcelReader(String fileName) {

        this.fileName = fileName;

        ClassLoader classLoader = getClass().getClassLoader();
        URL resource = classLoader.getResource(fileName);
        System.out.println("filePath " + resource);


        if (resource == null) {
            String s = String.format("Заданного файла " + fileName + " не существует.");
            logger.error(s);
            throw new IllegalArgumentException(s);
        } else {

            try {
                workbook = WorkbookFactory.create(new File(resource.getFile()));
            } catch (FileNotFoundException e) {
                String s = String.format("Заданного файла " + fileName + " не существует.");
                logger.error(s);
                System.exit(0);
            } catch (IOException e) {
                e.printStackTrace();
                logger.error("Ошибка.");
            } catch (InvalidFormatException e) {
                e.printStackTrace();
                logger.error("Ошибка InvalidFormatException.");
            }
        }
    }


    /**
     * Функция поиска текста в эксель файле.
     *
     * @param text             - текст, который ищем в файле.
     * @param equalsOrContains - true - полное совпадение @param text c текстом в ячейке, false - частичное.
     */
    @Override
    public void findText(String text, boolean equalsOrContains) {
        textCount = 0;

        workbook.forEach(sheet -> {
            sheet.forEach(row -> {
                row.forEach(cell -> {
                    if (cell.getCellTypeEnum() == CellType.STRING) {
                        BiPredicate<String, String> predicate = equalsOrContains ? String::equals : String::contains;
                        if (predicate.test(cell.toString(), text)) {
                            // System.out.printf("Текст %s найден в листе %s в ячейке %s.\n", text, cell.getSheet().getSheetName(), cell.getAddress());
                            String s = String.format("Текст %s найден в листе %s в ячейке %s.", text, cell.getSheet().getSheetName(), cell.getAddress());
                            System.out.println(s);
                            textCount++;
                        }
                    }
                });
            });
        });

        if (textCount != 0) {
            String s = String.format("Искомый текст " + text + " встречается в файле " + textCount + " раз(а).");
            logger.info(s);
        } else if (textCount == 0) {
            String s = String.format("Искомого текста " + text + " в файле не найдено.");
            logger.info(s);
        }
        try {
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Override
    public void insertImage(String fileName, String imageName, String sheetName, String cell) {
        listCount = 0;

        workbook.forEach(sheet -> {
            if (sheet.getSheetName().equals(sheetName)) {
                listCount++;
            }
        });


        System.out.println(listCount);
        if (listCount == 0) {
            String s = String.format("Листа %s не найдено. Создаем новый лист", sheetName);
            logger.info(s);


            String safeName = WorkbookUtil.createSafeSheetName(sheetName);
            workbook.createSheet(safeName);
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

    private void createList(String sheetName) {
        listCount = 0;

        workbook.forEach(sheet -> {
            if (sheet.getSheetName().equals(sheetName)) {
                listCount++;
            }
        });

        if (listCount == 0) {
            String s = String.format("Листа %s не найдено. Создаем новый лист", sheetName);
            logger.info(s);
            workbook.createSheet(sheetName);
        }

        FileOutputStream fileOut = null;
        try {
            fileOut = new FileOutputStream(new File(fileName));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        try {
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }
        try {
            fileOut.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    public void addImage(String sheetName, String imageName) {

        createList(sheetName);


        ClassLoader classLoader = getClass().getClassLoader();
        URL resource = classLoader.getResource(imageName);
        InputStream inputStream = null;
        try {
            inputStream = new FileInputStream(new File(resource.getFile()));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }


        byte[] bytes = new byte[0];
        try {
            bytes = IOUtils.toByteArray(inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        //Adds a picture to the workbook
        int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
        //close the input stream
        try {
            inputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        //Returns an object that handles instantiating concrete classes
        CreationHelper helper = workbook.getCreationHelper();
        //Creates the top-level drawing patriarch.

        final Drawing[] drawing = new Drawing[1];
        workbook.forEach(sheet -> {
            if (sheet.getSheetName().equals(sheetName)) {
              drawing[0] = sheet.createDrawingPatriarch();
            }
        });


        //Create an anchor that is attached to the worksheet
        ClientAnchor anchor = helper.createClientAnchor();

        //create an anchor with upper left cell _and_ bottom right cell
        anchor.setCol1(1); //Column B
        anchor.setRow1(2); //Row 3
        anchor.setCol2(2); //Column C
        anchor.setRow2(3); //Row 4

        //Creates a picture
        Picture pict = drawing[0].createPicture(anchor, pictureIdx);

        //Reset the image to the original size
        //pict.resize(); //don't do that. Let the anchor resize the image!

        //Create the Cell B3
        workbook.forEach(sheet -> {
            if (sheet.getSheetName().equals(sheetName)) {
                Cell cell = sheet.createRow(2).createCell(1);
            }
        });


        //set width to n character widths = count characters * 256
        //int widthUnits = 20*256;
        //sheet.setColumnWidth(1, widthUnits);

        //set height to n points in twips = n * 20
        //short heightUnits = 60*20;
        //cell.getRow().setHeight(heightUnits);

        //Write the Excel file
        FileOutputStream fileOut = null;
        try {
            fileOut = new FileOutputStream(new File(fileName));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        try {
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }
        try {
            fileOut.close();
        } catch (IOException e) {
            e.printStackTrace();
        }


    }
}

