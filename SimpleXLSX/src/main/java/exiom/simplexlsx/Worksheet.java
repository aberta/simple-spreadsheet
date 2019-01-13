package exiom.simplexlsx;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class Worksheet {

    private Workbook workbook;
    private String name;
    private int number;
    
    private List<Row> rows;

    Worksheet(Workbook workbook, String name) {
        this.workbook = workbook;
        this.name = name;
        rows = new ArrayList<Row>();
    }
    
    public Row nextRow() {
        return newRow(rows.size() + 1);
    }
    
    public Row newRow(int rowNum) {
        Row r = new Row(this, rowNum);
        rows.add(r);
        return r;        
    }
    
    void write(ZipOutputStream out) throws IOException, ParserConfigurationException, TransformerConfigurationException, TransformerException {

        out.putNextEntry(new ZipEntry("xl/worksheets/sheet" + getSheetNumber() + ".xml"));

        DocumentBuilder db = DocumentBuilderFactory.newInstance().newDocumentBuilder();
        Document doc = db.newDocument();

        Element root = doc.createElementNS("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "worksheet");
        root.setAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        doc.appendChild(root);

        Element sheetData = doc.createElement("sheetData");
        root.appendChild(sheetData);

        for (Row r: rows) {
            r.write(doc, sheetData);
        }
        
        XMLUtils.write(doc, out);

        out.closeEntry();
    }
    
    public String getName() {
        return name;
    }

    public int getSheetNumber() {
        return number;
    }

    void setNumber(int number) {
        this.number = number;
    }
    
    Workbook getWorkbook() {
        return workbook;
    }
}
