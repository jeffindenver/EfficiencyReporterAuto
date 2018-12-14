package excelops;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.StringJoiner;
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

    public ExcelOps() {

    }

    public XSSFWorkbook openWorkbook(String filename) throws InvalidFormatException, IOException {
        File file = new File(filename);
        XSSFWorkbook wb = new XSSFWorkbook();

        try {
            wb = XSSFWorkbookFactory.createWorkbook(file, true);
        } catch (InvalidFormatException | IOException e) {
            System.out.println(e.getMessage());
            throw e;
        } finally {
            if (wb != null) {
                try {
                    wb.close();
                } catch (IOException e) {
                    System.out.println(e.getMessage());
                    throw e;
                }
            }
        }
        return wb;
    }

    public List<String> sheetToList(XSSFWorkbook wb) {
        XSSFSheet sheet = wb.getSheetAt(0);
        return toListHelper(sheet);
    }

    public List<String> sheetToList(XSSFWorkbook wb, int sheetIndex) {
        XSSFSheet sheet = wb.getSheetAt(sheetIndex);
        return toListHelper(sheet);
    }

    private List<String> toListHelper(XSSFSheet sheet) {
        List<String> list = new ArrayList<>();
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
            list.add(joiner.toString());
        }
        return list;
    }

    public void writeWorkbook(XSSFWorkbook wb, String filename) {
        try (OutputStream fileOut = new FileOutputStream(filename)) {
            wb.write(fileOut);
        } catch (FileNotFoundException e) {
            System.out.println(e.getMessage());
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }
    }

    public void printList(List<String> list) {
        for (String s : list) {
            System.out.println(s);
        }
    }

}
