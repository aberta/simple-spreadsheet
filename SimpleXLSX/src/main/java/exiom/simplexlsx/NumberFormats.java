package exiom.simplexlsx;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class NumberFormats implements Iterable<NumberFormat> {

    private List<NumberFormat> formats;
    
    NumberFormats() {
        formats = new ArrayList<NumberFormat>();
    }
    
    public int getNumFmtId(String formatCode) {
    
        for (int i = 0; i < formats.size(); i++) {
            NumberFormat nf = formats.get(i);
            if (nf.getFormatCode().endsWith(formatCode)) {
                return nf.getNumFmtId();
            }
        }
        
        NumberFormat nf = new NumberFormat(formats.size() + 1, formatCode);
        formats.add(nf);
        return nf.getNumFmtId();
    }
    
    public int getNumberOfFormats() {
        return formats.size();
    }
    
    public NumberFormat get(int numFmtId) {
        for (NumberFormat nf: formats) {
            if (nf.getNumFmtId() == numFmtId) {
                return nf;
            }
        }
        return null;
    }
    
    public void write(Document doc, Element parent) {
        
        Element numFmts = doc.createElement("numFmts");
        parent.appendChild(numFmts);
        numFmts.setAttribute("count", Integer.toString(getNumberOfFormats()));

        int i = 0;
        for (NumberFormat nf: formats) {
            nf.write(doc, numFmts);
        }
    }

    @Override
    public Iterator<NumberFormat> iterator() {
        return formats.iterator();
    }
}
