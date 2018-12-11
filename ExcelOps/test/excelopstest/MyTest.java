package excelopstest;

import excelops.ExcelOps;
import java.io.IOException;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
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

    public MyTest() {
    }

    @Before
    public void setUp() {
        excelOps = new ExcelOps();
    }

    @After
    public void tearDown() {
    }

    @Test
    public void testOpenWorkbook() {
        XSSFWorkbook wb = null;

        try {
            wb = excelOps.openWorkbook("TestDummy.xlsx");
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
            wb = excelOps.openWorkbook("CellOne Nov 25.xlsx");
        } catch (InvalidFormatException | IOException e) {
            System.out.println(e.getMessage());
        }
        if (wb != null) {
            excelOps.setWorkbook(wb);
            list = excelOps.sheetToList(excelOps.getWorkbook());
            assertNotNull("Failed: list is null", list);
            excelOps.printList(list);
        }
    }
    
    @Test
    public void testIndexedSheetToList() {
        XSSFWorkbook wb = null;
        int sheetIndex = 1;
        try {
            wb = excelOps.openWorkbook("CellOne Nov 25.xlsx");
        } catch (InvalidFormatException | IOException e) {
            System.out.println(e.getMessage());
        }
        if (wb != null) {
            excelOps.setWorkbook(wb);
            list = excelOps.sheetToList(excelOps.getWorkbook(), sheetIndex);
            assertNotNull("Failed: list is null", list);
            excelOps.printList(list);
        }
    }

}
