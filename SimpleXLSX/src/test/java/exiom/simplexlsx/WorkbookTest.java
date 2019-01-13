/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package exiom.simplexlsx;

import java.io.IOException;
import java.math.BigDecimal;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.junit.Test;
import static org.junit.Assert.*;

public class WorkbookTest {
    
    public WorkbookTest() {
    }

    @Test
    public void testWorksheet() {
        Workbook wb = new Workbook();
        Worksheet ws = wb.addWorksheet("My Data");
        
        ws.nextRow()
                .append("Hello")
                .append(new BigDecimal("100.00"))
                .append(new Date())
                .append("001");
        ws.newRow(3)
                .append("World")
                .append(new BigDecimal("0.0"))
                .append(new Date());
        
        ws = wb.addWorksheet("My Other Data");
        ws.nextRow()
                .append("Hello World")
                .skipColumn()
                .append(new BigDecimal("9.900"))
                .append(new Date());
        
        try {
            wb.write("test.xlsx");
        } catch (IOException ex) {
            Logger.getLogger(WorkbookTest.class.getName()).log(Level.SEVERE, null, ex);
        }
        
        assertTrue(true);        
    }

}
