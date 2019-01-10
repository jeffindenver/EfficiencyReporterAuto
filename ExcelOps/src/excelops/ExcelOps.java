package excelops;

import java.io.File;
import java.io.FileInputStream;
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
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author JShepherd
 */
public class ExcelOps {

    public ExcelOps() {
    }

    public Workbook openWorkbook(String filename) throws InvalidFormatException, IOException {
        File file = new File(filename);
        if (file.exists()) {
            try (FileInputStream fis = new FileInputStream(filename)) {
                return WorkbookFactory.create(fis);
            }
        } else {
            return new XSSFWorkbook();
        }
    }

    public List<String> sheetToTabList(Workbook wb, int sheetIndex) {
        Sheet sheet = wb.getSheetAt(sheetIndex);
        String delim = "\t";
        return toListHelper(sheet, delim);
    }
    
    public List<String> sheetToList(Workbook wb, int sheetIndex) {
        Sheet sheet = wb.getSheetAt(sheetIndex);
        String delim = ",";
        return toListHelper(sheet, delim);
    }

    private List<String> toListHelper(Sheet sheet, String delim) {
        List<String> list = new ArrayList<>();
        Iterator<Row> rowIterator = sheet.iterator();

        while (rowIterator.hasNext()) {
            Row currentRow = rowIterator.next();
            Iterator<Cell> cellIterator = currentRow.iterator();
            StringJoiner joiner = new StringJoiner(delim);

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
