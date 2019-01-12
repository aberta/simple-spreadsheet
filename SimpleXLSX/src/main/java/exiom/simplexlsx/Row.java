package exiom.simplexlsx;

import java.util.ArrayList;
import java.util.List;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class Row {

    private int rowNum;    
    private List<Cell> cells;
    
    public Row(int rowNum) {
        this.rowNum = rowNum;
        cells = new ArrayList<Cell>();
    }
    
    public Row append(Cell cell) {
        cells.add(cell);
        return this;
    }
    
    public void write(Document doc, Element parent) {
        Element row = doc.createElement("row");
        row.setAttribute("r", Integer.toString(rowNum));
        parent.appendChild(row);        
        
        for (Cell c: cells) {
            c.write(doc, row);
        }
    }
}
