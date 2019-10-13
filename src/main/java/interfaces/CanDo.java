package interfaces;

public interface CanDo {
    void findText(String text, boolean equalsOrContains);
    void insertImage(String fileName, String imageName, String sheetName, String cell);
}
