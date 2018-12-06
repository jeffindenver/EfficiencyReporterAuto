package excelopstest;

import excelops.ExcelOps;
import java.util.List;
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
    private String filename;
    
    public MyTest() {
    }
    
    @Before
    public void setUp() {
        filename = "CellOne Nov 25.xlsx";
        excelOps = new ExcelOps(filename);
        list = excelOps.getList();
    }
    
    @After
    public void tearDown() {
    }

    @Test
    public void testReadWorkbook() {
        assertEquals("CellOne Nov 25.xlsx", filename);
        assertNotNull("Failed: excelOps is null", excelOps);        
    }
    
    @Test
    public void testWorkbookToList() {
        assertNotNull("Failed: list is null", list);
        excelOps.printList(list);

    }

}
