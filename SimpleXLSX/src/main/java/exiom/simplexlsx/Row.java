package exiom.simplexlsx;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class Row {

    private Worksheet sheet;
    private int rowNum;
    private List<Cell> cells;

    Row(Worksheet sheet, int rowNum) {
        this.sheet = sheet;
        this.rowNum = rowNum;
        cells = new ArrayList<Cell>();
    }

    public Row append(String text) {
        if (!getWorksheet().shouldIgnore(text)) {
            append(new Cell(this, text));
        }
        return this;
    }

    public Row append(BigDecimal number) {
        append(new Cell(this, number));
        return this;
    }

    public Row append(Date date) {
        append(new Cell(this, date));
        return this;
    }

    Row append(Cell cell) {
        cells.add(cell);
        return this;
    }
    
    public Row skipColumn() {
        return append("");
    }
    
    Worksheet getWorksheet() {
        return sheet;
    }
    
    public int getRowNum() {
        return rowNum;
    }

    public void write(Document doc, Element parent) {
        Element row = doc.createElement("row");
        row.setAttribute("r", Integer.toString(rowNum));
        parent.appendChild(row);

        for (Cell c : cells) {
            c.write(doc, row);
        }
    }
}
