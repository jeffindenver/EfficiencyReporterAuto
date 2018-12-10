package excelops;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.StringJoiner;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory;

/**
 *
 * @author JShepherd
 */
public class ExcelOps {

    private final List<String> list;
    private XSSFWorkbook workbook;

    public ExcelOps() {
        list = new ArrayList<>();

    }

    public XSSFWorkbook openWorkbook(String filename) throws InvalidFormatException, IOException {
        //if the workbook does not already exist, a new workbook is created.
        File file = new File(filename);
        XSSFWorkbook wb = null;

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

    public void toList(XSSFWorkbook wb) {
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

    private void writeSheet() {
        int endSheet = workbook.getNumberOfSheets();
        XSSFSheet newSheet = workbook.cloneSheet(endSheet);
        XSSFRow row = newSheet.getRow(5);
        //todo clear contents of all rows in separate method
        //write new data to correct range. Add row range an argument to this
        //method.
    }

    public void printList(List<String> list) {
        for (String s : list) {
            System.out.println(s);
        }
    }

    public List<String> getList() {
        return new ArrayList<>(list);
    }

    public void setWorkbook(XSSFWorkbook wb) {
        workbook = wb;
    }

    public XSSFWorkbook getWorkbook() {
        return workbook;
    }

}
