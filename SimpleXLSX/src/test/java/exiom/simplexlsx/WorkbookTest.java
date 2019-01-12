/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package exiom.simplexlsx;

import java.io.OutputStream;
import org.junit.Test;
import static org.junit.Assert.*;

/**
 *
 * @author chris
 */
public class WorkbookTest {
    
    public WorkbookTest() {
    }

    @Test
    public void testWorksheet() {
        Workbook wb = new Workbook();
        wb.addWorksheet("My Data");
        wb.addWorksheet("My Other Data");
        
        wb.write((OutputStream)null);
        
        assertTrue(true);        
    }

}
