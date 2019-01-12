package exiom.simplexlsx;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class Workbook {

    private List<Worksheet> worksheets;

    public Workbook() {
        worksheets = new ArrayList<Worksheet>();
    }

    public Worksheet addWorksheet(String name) {
        Worksheet ws = new Worksheet(name);
        worksheets.add(ws);
        return ws;
    }

    public void write(String filename) throws IOException {
        ZipOutputStream out = new ZipOutputStream(new FileOutputStream(filename));
        write(out);
    }

    public void write(ZipOutputStream out) throws IOException {

        int i = 1;
        for (Worksheet ws : worksheets) {
            ws.setNumber(i++);
        }

        try {
            writeRels(out);
            writeContentTypes(out);
            writeWorkbook(out);
            out.close();
        } catch (ParserConfigurationException ex) {
            throw new RuntimeException(ex);
        } catch (TransformerException ex) {
            throw new RuntimeException(ex);
        }
    }

    private void writeContentTypes(ZipOutputStream out) throws ParserConfigurationException,
            TransformerConfigurationException, TransformerException, IOException {

        out.putNextEntry(new ZipEntry("[Content_Types].xml"));

        Document doc = XMLUtils.newDocument("http://schemas.openxmlformats.org/package/2006/content-types", "Types");
        Element root = doc.getDocumentElement();

        Element dflt = doc.createElement("Default");
        dflt.setAttribute("Extension", "rels");
        dflt.setAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml");
        root.appendChild(dflt);

        Element override = doc.createElement("Override");
        override.setAttribute("PartName", "/xl/_rels/workbook.xml.rels");
        override.setAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml");
        root.appendChild(override);

        override = doc.createElement("Override");
        override.setAttribute("PartName", "/xl/workbook.xml");
        override.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
        root.appendChild(override);

        override = doc.createElement("Override");
        override.setAttribute("PartName", "/xl/worksheets/sheet1.xml");
        override.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
        root.appendChild(override);

        override = doc.createElement("Override");
        override.setAttribute("PartName", "/xl/worksheets/sheet2.xml");
        override.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
        root.appendChild(override);

        XMLUtils.write(doc, out);

        out.closeEntry();
    }

    private void writeRels(ZipOutputStream out) throws ParserConfigurationException,
            TransformerConfigurationException, TransformerException, IOException {

        writeRootRels(out);
        writeWorkbookRels(out);
    }

    private void writeRootRels(ZipOutputStream out) throws ParserConfigurationException,
            TransformerConfigurationException, TransformerException, IOException {

        out.putNextEntry(new ZipEntry("_rels/.rels"));

        Document doc = XMLUtils.newDocument("http://schemas.openxmlformats.org/package/2006/relationships", "Relationships");
        Element root = doc.getDocumentElement();

        Element rel = doc.createElement("Relationship");
        rel.setAttribute("Id", "rId1");
        rel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
        rel.setAttribute("Target", "/xl/workbook.xml");
        root.appendChild(rel);

        XMLUtils.write(doc, out);

        out.closeEntry();
    }

    private void writeWorkbookRels(ZipOutputStream out) throws ParserConfigurationException,
            TransformerConfigurationException, TransformerException, IOException {

        out.putNextEntry(new ZipEntry("xl/_rels/workbook.xml.rels"));

        Document doc = XMLUtils.newDocument("http://schemas.openxmlformats.org/package/2006/relationships", "Relationships");
        Element root = doc.getDocumentElement();

        for (Worksheet ws : worksheets) {
            Element rel = doc.createElement("Relationship");
            rel.setAttribute("Id", "rId" + Integer.toString(ws.getSheetNumber() + 1));
            rel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
            rel.setAttribute("Target", "/xl/worksheets/sheet" + ws.getSheetNumber()+ ".xml");
            root.appendChild(rel);
        }

        XMLUtils.write(doc, out);

        out.closeEntry();
    }

    private void writeWorkbook(ZipOutputStream out) throws ParserConfigurationException,
            TransformerConfigurationException, TransformerException, IOException {

        out.putNextEntry(new ZipEntry("xl/workbook.xml"));

        Document doc = XMLUtils.newDocument("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "workbook");
        Element root = doc.getDocumentElement();
        root.setAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        Element sheets = doc.createElement("sheets");
        root.appendChild(sheets);

        for (Worksheet ws : worksheets) {
            Element sheet = doc.createElement("sheet");
            sheet.setAttribute("name", ws.getName());
            sheet.setAttribute("sheetId", Integer.toString(ws.getSheetNumber()));
            sheet.setAttribute("r:id", "rId" + Integer.toString(ws.getSheetNumber() + 1));
            sheets.appendChild(sheet);
        }

        XMLUtils.write(doc, out);

        out.closeEntry();

        for (Worksheet ws : worksheets) {
            ws.write(out);
        }
    }
}
