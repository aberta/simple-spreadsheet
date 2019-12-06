package exiom.simplexlsx;

import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class Cell {

    private Row row;
    private Object value;

    public Cell(Row row, String text) {
        this.row = row;
        this.value = text;
    }

    public Cell(Row row, BigDecimal number) {
        this.row = row;
        value = number;
    }

    public Cell(Row row, Date date) {
        this.row = row;
        value = date;
    }

    void write(Document doc, Element parent) {

        Element cell = doc.createElement("c");
        parent.appendChild(cell);

        if (value instanceof String) {

            cell.setAttribute("t", "inlineStr");
            cell.appendChild(doc.createElement("is")).appendChild(doc.createElement("t"))
                    .appendChild(doc.createTextNode((String) value));

        } else if (value instanceof BigDecimal) {
            BigDecimal number = (BigDecimal) value;
            cell.setAttribute("t", "n");
            cell.appendChild(doc.createElement("v")).appendChild(doc.createTextNode(number.toPlainString()));

            if (number.scale() > 0) {
                int styleId = row.getWorksheet().getWorkbook().getStyleForNumFmt(number.scale());
                cell.setAttribute("s", Integer.toString(styleId));
            }
        } else if (value instanceof Date) {

            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss");
            String iso8601 = sdf.format((Date) value);

            cell.setAttribute("t", "d");
            cell.appendChild(doc.createElement("v")).appendChild(doc.createTextNode(iso8601));

            int styleId = row.getWorksheet().getWorkbook().getStyleForDate();
            cell.setAttribute("s", Integer.toString(styleId));
        }
    }
}
