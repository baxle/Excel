package Interface;

public interface CanDo {
    void findText(String fileName, String text, boolean equalsOrContains);
    void insertImage(String fileName, String imageName, String sheetName, String cell);
}
