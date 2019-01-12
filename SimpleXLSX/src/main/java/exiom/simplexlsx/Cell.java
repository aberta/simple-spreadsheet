package exiom.simplexlsx;

import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class Cell {

    private Object value;

    public Cell(String text) {
        this.value = text;
    }

    public Cell(BigDecimal number) {
        value = number;
    }

    public Cell(Date date) {
        value = date;
    }

    void write(Document doc, Element row) {

        Element cell = doc.createElement("c");
        row.appendChild(cell);

        if (value instanceof String) {

            cell.setAttribute("t", "inlineStr");
            cell.appendChild(doc.createElement("is"))
                    .appendChild(doc.createElement("t"))
                    .appendChild(doc.createTextNode((String) value));

        } else if (value instanceof BigDecimal) {
            cell.setAttribute("t", "n");
            cell.appendChild(doc.createElement("v"))
                    .appendChild(doc.createTextNode(((BigDecimal) value).toPlainString()));

        } else if (value instanceof Date) {

            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss");
            String iso8601 = sdf.format((Date) value);

            cell.setAttribute("t", "d");
            cell.appendChild(doc.createElement("v")).appendChild(doc.createTextNode(iso8601));
        }
    }
}
