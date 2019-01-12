package exiom.simplexlsx;

import java.io.OutputStream;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
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

public class XMLUtils {

    public static Document newDocument(String ns, String rootTag) {

        DocumentBuilder db;
        try {
            db = DocumentBuilderFactory.newInstance().newDocumentBuilder();
            Document doc = db.newDocument();

            Element root = doc.createElementNS(ns, rootTag);
            doc.appendChild(root);

            return doc;
        } catch (ParserConfigurationException ex) {
            throw new RuntimeException(ex);
        }
    }

    public static void write(Document doc, OutputStream out) {
        try {
            Transformer t = TransformerFactory.newInstance().newTransformer();
            t.setOutputProperty(OutputKeys.INDENT, "yes");
            t.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "2");

            Source src = new DOMSource(doc);
            Result dest = new StreamResult(out);

            t.transform(src, dest);

        } catch (TransformerConfigurationException ex) {
            throw new RuntimeException(ex);
        } catch (TransformerException ex) {
            throw new RuntimeException(ex);
        }
    }
}
