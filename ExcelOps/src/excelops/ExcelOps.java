package excelops;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.StringJoiner;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory;

/**
 *
 * @author JShepherd
 */
public class ExcelOps {

    private final List<String> list;

    public ExcelOps(String filename) {
        list = new ArrayList<>();
        readWorkbook(filename);
    }

    private void readWorkbook(String filename) {
        XSSFWorkbook wb = null;
        File file = new File(filename);

        try {
            wb = XSSFWorkbookFactory.createWorkbook(file, true);
            toList(wb);
        } catch (InvalidFormatException | IOException | EncryptedDocumentException e) {
            System.out.println(e.getMessage());
        } finally {
            if (wb != null) {
                try {
                    wb.close();
                } catch (IOException e) {
                    System.out.println(e.getMessage());
                }
            }
        }
    }

    private void toList(XSSFWorkbook wb) {
        XSSFSheet sheet = wb.getSheetAt(0);
        Iterator<Row> rowIterator = sheet.iterator();

        while (rowIterator.hasNext()) {
            Row currentRow = rowIterator.next();
            Iterator<Cell> cellIterator = currentRow.iterator();
            StringJoiner joiner = new StringJoiner(",");

            while (cellIterator.hasNext()) {
                Cell currentCell = cellIterator.next();
                if (currentCell.getCellType() == CellType.STRING) {
                    joiner.add(currentCell.getStringCellValue());
                } else if (currentCell.getCellType() == CellType.NUMERIC) {
                    joiner.add(String.valueOf(currentCell.getNumericCellValue()));
                }
            }
            System.out.println(joiner.toString());
            list.add(joiner.toString());
        }
    }

    public void printList(List<String> list) {
        for (String s : list) {
            System.out.println(s);
        }
    }

    public List<String> getList() {
        return new ArrayList<>(list);
    }


}
