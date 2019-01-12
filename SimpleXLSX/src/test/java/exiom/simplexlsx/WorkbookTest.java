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
        
        ws.add(new Row(1))
                .append(new Cell("Hello"))
                .append(new Cell(new BigDecimal("100.00")))
                .append(new Cell(new Date()));
        ws.add(new Row(2))
                .append(new Cell("World"))
                .append(new Cell(new BigDecimal("0.0")))
                .append(new Cell(new Date()));
        
        ws = wb.addWorksheet("My Other Data");
        ws.add(new Row(1))
                .append(new Cell("Hello World"))
                .append(new Cell(new BigDecimal("9.9")))
                .append(new Cell(new Date()));
        
        try {
            wb.write("test.xlsx");
        } catch (IOException ex) {
            Logger.getLogger(WorkbookTest.class.getName()).log(Level.SEVERE, null, ex);
        }
        
        assertTrue(true);        
    }

}
