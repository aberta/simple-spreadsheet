package exiom.simplexlsx;

import java.io.File;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Result;
import javax.xml.transform.Source;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class Workbook {

    private Map<String, Worksheet> worksheets;

    public Workbook() {
        worksheets = new HashMap<String, Worksheet>();
    }

    public Worksheet addWorksheet(String name) {
        return worksheets.put(name, new Worksheet(name));
    }

    public Worksheet getWorksheet(String name) {
        return worksheets.get(name);
    }

    public void write(String filename) {

    }

    public void write(OutputStream out) {
        try {
            writeWorkbook();
        } catch (ParserConfigurationException ex) {
            Logger.getLogger(Workbook.class.getName()).log(Level.SEVERE, null, ex);
        } catch (TransformerException ex) {
            Logger.getLogger(Workbook.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    private void writeWorkbook() throws ParserConfigurationException, TransformerConfigurationException, TransformerException {
        /*
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <sheets>
        <sheet name="Brian" sheetId="1" r:id="rId1"/>
        </sheets>
        </workbook>
         */
        DocumentBuilder db = DocumentBuilderFactory.newInstance().newDocumentBuilder();
        Document doc = db.newDocument();

        Element root = doc.createElementNS("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "workbook");
        root.setAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        doc.appendChild(root);
        
        Element sheets = doc.createElement("sheets");
        root.appendChild(sheets);
        
        int sheetId = 0;
        for (Worksheet ws: worksheets.values()) {
            sheetId++;
            Element sheet = doc.createElement("sheet");
            sheet.setAttribute("name", ws.getName());
            sheet.setAttribute("sheetId", Integer.toString(sheetId));
            sheet.setAttribute("r:id", "rId" + sheetId);
            sheets.appendChild(sheet);
        }

        Transformer t = TransformerFactory.newInstance().newTransformer();
        Source src = new DOMSource(doc);
        Result dest = new StreamResult(new File("workbook.xml"));
        t.transform(src, dest);
        
        
    }    
}
