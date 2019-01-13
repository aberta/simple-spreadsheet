package exiom.simplexlsx;

import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class NumberFormat {

    private static String DATE_FORMAT = "yyyy\\-mm\\-dd;@";
        
    private int id;
    private String formatCode;

    public static NumberFormat forNumber(int id, int decimalPositions) {
        return new NumberFormat(id, getFormatCodeForDecimals(decimalPositions));        
    }

    public static NumberFormat forDate(int id) {
        return new NumberFormat(id, DATE_FORMAT);        
    }
    
    NumberFormat(int id, String formatCode) {
        this.id = id;
        this.formatCode = formatCode;
    }
    
    public static String getFormatCodeForDecimals(int decimalPositions) {
        StringBuilder sb = new StringBuilder();
        sb.append("0.");
        for (int i = 0; i < decimalPositions; i++) {
            sb.append("0");
        }
        return sb.toString();
    }
    
    public static String getFormatForDate() {
        return DATE_FORMAT;
    }

    public int getNumFmtId() {
        return id;
    }
    
    public String getFormatCode() {
        return formatCode;
    }

    public void write(Document doc, Element parent) {
        Element fmt = doc.createElement("numFmt");
        parent.appendChild(fmt);
        fmt.setAttribute("numFmtId", Integer.toString(getNumFmtId()));
        fmt.setAttribute("formatCode", getFormatCode());
    }
}
