package excelopstest;

import excelops.ExcelOps;
import java.io.File;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.After;
import org.junit.AfterClass;
import static org.junit.Assert.*;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;

/**
 *
 * @author JShepherd
 */
public class MyTest {

    @BeforeClass
    public static void setUpClass() {
    }

    @AfterClass
    public static void tearDownClass() {
    }

    private ExcelOps excelOps;
    private List<String> list;
    private String testFilename;
    private String testHSSFFilename;

    public MyTest() {
    }

    @Before
    public void setUp() {
        excelOps = new ExcelOps();
        testFilename = "CellOne Nov 25.xlsx";
        testHSSFFilename = "CellOne Nov 25.xls";
    }

    @After
    public void tearDown() {
    }

    @Test
    public void testOpenWorkbook() {
        XSSFWorkbook wb = null;

        try {
            wb = (XSSFWorkbook) excelOps.openWorkbook("TestDummy.xlsx");
        } catch (InvalidFormatException | IOException e) {
            System.out.println(e.getMessage());
        }

        assertNotNull("Failed: workbook is null", wb);
        System.out.println("Before wb.toString()"
                + "\n"
                + wb.toString()
                + "\n"
                + "After wb.toString()");
    }

    @Test
    public void testSheetToList() {
        XSSFWorkbook wb = null;
        try {
            wb = (XSSFWorkbook) excelOps.openWorkbook(testFilename);
        } catch (InvalidFormatException | IOException e) {
            System.out.println(e.getMessage());
        }
        if (wb != null) {
            list = excelOps.sheetToList(wb, 0);
            assertNotNull("Failed: list is null", list);
            excelOps.printList(list);
        }
    }

    @Test
    public void testIndexedHSSFSheetToList() {
        HSSFWorkbook wb = null;
        int sheetIndex = 0;
        try {
            wb = (HSSFWorkbook) excelOps.openWorkbook(testHSSFFilename);
        } catch (InvalidFormatException | IOException e) {
            System.out.println(e.getMessage());
        }
        if (wb != null) {
            list = excelOps.sheetToList(wb, sheetIndex);
            assertNotNull("Failed: list is null", list);
            excelOps.printList(list);
        }
    }
    
    @Test
    public void testIndexedSheetToList() {
        XSSFWorkbook wb = null;
        int sheetIndex = 1;
        try {
            wb = (XSSFWorkbook) excelOps.openWorkbook(testFilename);
        } catch (InvalidFormatException | IOException e) {
            System.out.println(e.getMessage());
        }
        if (wb != null) {
            list = excelOps.sheetToList(wb, sheetIndex);
            assertNotNull("Failed: list is null", list);
            excelOps.printList(list);
        }
    }

    @Test
    public void testWriteWorkbook() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");

        XSSFRow row = sheet.createRow(0);
        List<String> v = Arrays.asList("Hello", "World", "This", "Is", "a", "test");

        for (int i = 0; i < v.size(); i++) {
            XSSFCell cell = row.createCell(i);
            cell.setCellValue(v.get(i));
        }

        excelOps.writeWorkbook(wb, "testWrite.xlsx");
    }
}
