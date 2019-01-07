package excelopstest;

import excelops.ExcelOps;
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
    private List<String> cpatList;
    private String testFilename;
    private String testHSSFFilename;

    public MyTest() {
    }

    @Before
    public void setUp() {
        excelOps = new ExcelOps();
        testFilename = "CellOne Nov 25.xlsx";
        testHSSFFilename = "CPaT Summary.xls";
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
            cpatList = excelOps.sheetToList(wb, sheetIndex);
            assertNotNull("Failed: list is null", cpatList);
            excelOps.printList(cpatList);
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
    public void testGetSummaryLine() {
        HSSFWorkbook wb = null;
        int sheetIndex = 0;
        try {
            wb = (HSSFWorkbook) excelOps.openWorkbook(testHSSFFilename);
        } catch (InvalidFormatException | IOException e) {
            System.out.println(e.getMessage());
        }
        if (wb != null) {
            cpatList = excelOps.sheetToList(wb, sheetIndex);
            assertNotNull("Failed: list is null", cpatList);
            excelOps.printList(cpatList);
        }
        
        String summary = "";

        int max = cpatList.size() - 1;
        
        for(int i = max; i >= 0; i--) {
            String[] elements = cpatList.get(i).split(",");
            if (elements[0].equalsIgnoreCase("Summary")) {
                summary = cpatList.get(i);
                break;
            }
        }
        System.out.println(summary);
        
        XSSFWorkbook xssfWb = new XSSFWorkbook();
        XSSFSheet sheet = xssfWb.createSheet("Summary");

        XSSFRow row = sheet.createRow(0);
        String[] v = summary.split(",");
        for (int i = 0; i < v.length; i++) {
            XSSFCell cell = row.createCell(i);
            cell.setCellValue(v[i]);
        }

        excelOps.writeWorkbook(xssfWb, "Summary.xlsx");
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
