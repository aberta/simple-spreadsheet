package exiom.simplexlsx;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
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

    public Row append(String text) {
        append(new Cell(text));
        return this;
    }

    public Row append(BigDecimal number) {
        append(new Cell(number));
        return this;
    }

    public Row append(Date date) {
        append(new Cell(date));
        return this;
    }

    public Row append(Cell cell) {
        cells.add(cell);
        return this;
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
