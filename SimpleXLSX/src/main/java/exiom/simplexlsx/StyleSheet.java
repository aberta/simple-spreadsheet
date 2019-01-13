package exiom.simplexlsx;

import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipOutputStream;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class StyleSheet {

    private List<Style> styles;
    private NumberFormats numFmts;

    public StyleSheet() {
        styles = new ArrayList<Style>();
        numFmts = new NumberFormats();
    }

    int getStyleForFormatCode(String formatCode) {

        for (int i = 0; i < styles.size(); i++) {
            Style s = styles.get(i);
            if (s.getNumFmtId() > 0) {
                NumberFormat nf = numFmts.get(s.getNumFmtId());
                if (nf != null) {
                    if (nf.getFormatCode().equals(formatCode)) {
                        return i + 1;
                    }
                }
            }
        }

        return addStyle(formatCode);
    }
    
    private int addStyle(String formatCode) {
        int numFmtId = numFmts.getNumFmtId(formatCode);
        Style s = new Style(numFmtId);
        styles.add(s);
        return styles.size();
    }

    void write(ZipOutputStream out) {

        Document doc = XMLUtils.newDocument("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "styleSheet");
        Element root = doc.getDocumentElement();

        numFmts.write(doc, root);

        appendEmptyElement(doc, root, "font");
        appendEmptyElement(doc, root, "fill");
        appendEmptyElement(doc, root, "border");

        Element cellStyleXfs = doc.createElement("cellStyleXfs");
        root.appendChild(cellStyleXfs);
        cellStyleXfs.setAttribute("count", Integer.toString(styles.size() + 1));

        Element xf = doc.createElement("xf");
        xf.setAttribute("numFmtId", "0");
        xf.setAttribute("fontId", "0");
        xf.setAttribute("fillId", "0");
        xf.setAttribute("borderId", "0");
        cellStyleXfs.appendChild(xf);

        Element cellXfs = doc.createElement("cellXfs");
        root.appendChild(cellXfs);
        cellXfs.setAttribute("count", Integer.toString(styles.size() + 1));

        xf = doc.createElement("xf");
        cellXfs.appendChild(xf);

        for (Style s : styles) {
            xf = doc.createElement("xf");
            xf.setAttribute("numFmtId", Integer.toString(s.getNumFmtId()));
            xf.setAttribute("fontId", "0");
            xf.setAttribute("fillId", "0");
            xf.setAttribute("borderId", "0");
            xf.setAttribute("xfId", "0");
            xf.setAttribute("applyNumberFormat", "1");
            cellXfs.appendChild(xf);
        }

        XMLUtils.write(doc, out);
    }

    private void appendEmptyElement(Document doc, Element parent, String elementName) {
        Element e = doc.createElement(elementName + "s");
        e.setAttribute("count", "1");
        e.appendChild(doc.createElement(elementName));
        parent.appendChild(e);
    }
}
