package vn.vietdoc.vietdoc_assistant;

import java.io.IOException;
import java.nio.file.*;
import java.net.URI;
import java.util.HashMap;
import java.util.Map;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

public class StyleSwapper {

    /**
     * Thay thế và cập nhật file styles.xml bên trong DOCX với các tham số tùy chỉnh
     * 
     * @param docxPath        Đường dẫn đến file DOCX
     * @param newStyleXmlPath Đường dẫn đến file styles_config.xml template
     * @param params          FormattingParameters chứa các tham số tùy chỉnh
     */
    public static void swapStyles(String docxPath, String newStyleXmlPath, FormattingParameters params) {
        System.out.println(
                ">>> [ZIP SWAP] Đang thay thế file styles.xml bên trong DOCX với params: " + params.toString());

        // Cấu hình để mở file Zip (docx thực chất là zip)
        Map<String, String> env = new HashMap<>();
        env.put("create", "false");

        // Chuyển đường dẫn file thành URI
        Path path = Paths.get(docxPath);
        URI uri = URI.create("jar:" + path.toUri());

        try (FileSystem fs = FileSystems.newFileSystem(uri, env)) {
            // 1. Đọc và parse file styles_config.xml từ bên ngoài
            Document styleDoc = parseXmlFile(newStyleXmlPath);

            // 2. Update các giá trị trong XML dựa vào FormattingParameters
            updateStylesXml(styleDoc, params);

            // 3. Ghi file XML đã update vào memory
            String updatedXmlContent = convertDocumentToString(styleDoc);
            byte[] xmlBytes = updatedXmlContent.getBytes("UTF-8");

            // 4. Xác định đường dẫn file styles.xml bên trong file Docx
            Path internalStylePath = fs.getPath("/word/styles.xml");

            // 5. Ghi XML đã update vào file styles.xml bên trong docx
            Files.write(internalStylePath, xmlBytes, StandardOpenOption.WRITE, StandardOpenOption.TRUNCATE_EXISTING);

            System.out.println(">>> [ZIP SWAP] Thành công! Đã cập nhật Style XML với các tham số tùy chỉnh vào file.");

        } catch (IOException e) {
            System.err.println("!!! LỖI KHI TRÁO STYLE: " + e.getMessage());
            e.printStackTrace();
        } catch (Exception e) {
            System.err.println("!!! LỖI KHI XỬ LÝ XML: " + e.getMessage());
            e.printStackTrace();
        }
    }

    /**
     * Parse file XML từ đường dẫn file
     */
    private static Document parseXmlFile(String filePath) throws Exception {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        factory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
        factory.setFeature("http://xml.org/sax/features/external-general-entities", false);
        factory.setFeature("http://xml.org/sax/features/external-parameter-entities", false);

        DocumentBuilder builder = factory.newDocumentBuilder();
        return builder.parse(new java.io.File(filePath));
    }

    /**
     * Update các giá trị trong XML dựa vào FormattingParameters
     * Cập nhật: font size, line spacing, spacing before/after, indentation
     */
    private static void updateStylesXml(Document doc, FormattingParameters params) {
        // 1. Update font sizes trong docDefaults
        updateDocDefaults(doc, params);

        // 2. Update từng style (Normal, Heading1, Heading2, etc.)
        updateStyleElement(doc, "Normal", params.getFontSizeBodyHalfPoint(), params);
        updateStyleElement(doc, "Heading1", params.getFontSizeHeading1HalfPoint(), params);
        updateStyleElement(doc, "Heading2", params.getFontSizeHeading2HalfPoint(), params);
        updateStyleElement(doc, "Heading3", params.getFontSizeHeading3HalfPoint(), params);

        // 3. Update caption styles nếu cần
        if (params.isItalicCaption()) {
            addItalicToStyle(doc, "BangStyle");
            addItalicToStyle(doc, "HinhStyle");
        }
    }

    /**
     * Update docDefaults - các thiết lập mặc định cho toàn bộ document
     */
    private static void updateDocDefaults(Document doc, FormattingParameters params) {
        NodeList docDefaultsList = doc.getElementsByTagName("w:docDefaults");
        if (docDefaultsList.getLength() > 0) {
            Element docDefaults = (Element) docDefaultsList.item(0);

            // Update trong rPrDefault (text properties)
            NodeList rPrDefaultList = docDefaults.getElementsByTagName("w:rPrDefault");
            if (rPrDefaultList.getLength() > 0) {
                Element rPrDefault = (Element) rPrDefaultList.item(0);
                NodeList rPrList = rPrDefault.getElementsByTagName("w:rPr");
                if (rPrList.getLength() > 0) {
                    Element rPr = (Element) rPrList.item(0);
                    updateFontSize(rPr, params.getFontSizeBodyHalfPoint());
                }
            }

            // Update trong pPrDefault (paragraph properties)
            NodeList pPrDefaultList = docDefaults.getElementsByTagName("w:pPrDefault");
            if (pPrDefaultList.getLength() > 0) {
                Element pPrDefault = (Element) pPrDefaultList.item(0);
                NodeList pPrList = pPrDefault.getElementsByTagName("w:pPr");
                if (pPrList.getLength() > 0) {
                    Element pPr = (Element) pPrList.item(0);
                    updateLineSpacing(pPr, params.getLineSpacingXmlUnits());
                }
            }
        }
    }

    /**
     * Update một style element cụ thể
     */
    private static void updateStyleElement(Document doc, String styleId, int fontSizeHalfPoint,
            FormattingParameters params) {
        NodeList styleList = doc.getElementsByTagName("w:style");

        for (int i = 0; i < styleList.getLength(); i++) {
            Element style = (Element) styleList.item(i);
            String id = style.getAttribute("w:styleId");

            if (id.equals(styleId)) {
                // Tìm hoặc tạo rPr (run properties)
                NodeList rPrList = style.getElementsByTagName("w:rPr");
                Element rPr = null;

                if (rPrList.getLength() > 0) {
                    rPr = (Element) rPrList.item(0);
                } else {
                    // Tạo rPr mới nếu chưa có
                    rPr = doc.createElement("w:rPr");

                    // Tìm pPr để insert rPr trước nó
                    NodeList pPrList = style.getElementsByTagName("w:pPr");
                    if (pPrList.getLength() > 0) {
                        style.insertBefore(rPr, pPrList.item(0));
                    } else {
                        style.appendChild(rPr);
                    }
                }

                // Update font size
                updateFontSize(rPr, fontSizeHalfPoint);

                // Tìm hoặc tạo pPr (paragraph properties)
                NodeList pPrList = style.getElementsByTagName("w:pPr");
                Element pPr = null;

                if (pPrList.getLength() > 0) {
                    pPr = (Element) pPrList.item(0);
                } else {
                    pPr = doc.createElement("w:pPr");
                    style.insertBefore(pPr, rPr);
                }

                // Update spacing và line spacing
                updateLineSpacing(pPr, params.getLineSpacingXmlUnits());
                updateParagraphSpacing(pPr, params.getSpacingBefore(), params.getSpacingAfter());

                break;
            }
        }
    }

    /**
     * Update font size trong element
     */
    private static void updateFontSize(Element element, int sizeHalfPoint) {
        // Tìm hoặc tạo w:sz element
        NodeList szList = element.getElementsByTagName("w:sz");
        Element sz = null;

        if (szList.getLength() > 0) {
            sz = (Element) szList.item(0);
        } else {
            sz = element.getOwnerDocument().createElement("w:sz");
            element.appendChild(sz);
        }

        sz.setAttribute("w:val", String.valueOf(sizeHalfPoint));

        // Cập nhật w:szCs (complex scripts) cùng với sz
        NodeList szCsList = element.getElementsByTagName("w:szCs");
        Element szCs = null;

        if (szCsList.getLength() > 0) {
            szCs = (Element) szCsList.item(0);
        } else {
            szCs = element.getOwnerDocument().createElement("w:szCs");
            element.appendChild(szCs);
        }

        szCs.setAttribute("w:val", String.valueOf(sizeHalfPoint));
    }

    /**
     * Update line spacing trong element
     */
    private static void updateLineSpacing(Element element, int lineSpacingXmlUnits) {
        // Tìm hoặc tạo w:spacing element
        NodeList spacingList = element.getElementsByTagName("w:spacing");
        Element spacing = null;

        if (spacingList.getLength() > 0) {
            spacing = (Element) spacingList.item(0);
        } else {
            spacing = element.getOwnerDocument().createElement("w:spacing");
            element.insertBefore(spacing, element.getFirstChild());
        }

        spacing.setAttribute("w:line", String.valueOf(lineSpacingXmlUnits));
        spacing.setAttribute("w:lineRule", "auto");
    }

    /**
     * Update paragraph spacing (before/after)
     */
    private static void updateParagraphSpacing(Element pPr, int spacingBefore, int spacingAfter) {
        NodeList spacingList = pPr.getElementsByTagName("w:spacing");
        Element spacing = null;

        if (spacingList.getLength() > 0) {
            spacing = (Element) spacingList.item(0);
        } else {
            spacing = pPr.getOwnerDocument().createElement("w:spacing");
            pPr.insertBefore(spacing, pPr.getFirstChild());
        }

        spacing.setAttribute("w:before", String.valueOf(spacingBefore));
        spacing.setAttribute("w:after", String.valueOf(spacingAfter));
    }

    /**
     * Thêm italic vào style element nếu chưa có
     */
    private static void addItalicToStyle(Document doc, String styleId) {
        NodeList styleList = doc.getElementsByTagName("w:style");

        for (int i = 0; i < styleList.getLength(); i++) {
            Element style = (Element) styleList.item(i);
            String id = style.getAttribute("w:styleId");

            if (id.equals(styleId)) {
                // Tìm hoặc tạo rPr
                NodeList rPrList = style.getElementsByTagName("w:rPr");
                Element rPr = null;

                if (rPrList.getLength() > 0) {
                    rPr = (Element) rPrList.item(0);
                } else {
                    rPr = doc.createElement("w:rPr");
                    NodeList pPrList = style.getElementsByTagName("w:pPr");
                    if (pPrList.getLength() > 0) {
                        style.insertBefore(rPr, pPrList.item(0));
                    } else {
                        style.appendChild(rPr);
                    }
                }

                // Kiểm tra có w:i chưa
                NodeList iList = rPr.getElementsByTagName("w:i");
                if (iList.getLength() == 0) {
                    Element italic = doc.createElement("w:i");
                    rPr.appendChild(italic);

                    // Cũng thêm w:iCs (complex scripts variant)
                    Element italicCs = doc.createElement("w:iCs");
                    rPr.appendChild(italicCs);
                }

                break;
            }
        }
    }

    /**
     * Chuyển đổi Document XML thành String
     */
    private static String convertDocumentToString(Document doc) throws Exception {
        TransformerFactory factory = TransformerFactory.newInstance();
        // factory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);

        Transformer transformer = factory.newTransformer();
        transformer.setOutputProperty("encoding", "UTF-8");
        transformer.setOutputProperty("version", "1.0");
        transformer.setOutputProperty("standalone", "yes");

        DOMSource source = new DOMSource(doc);
        java.io.StringWriter writer = new java.io.StringWriter();
        StreamResult result = new StreamResult(writer);

        transformer.transform(source, result);
        return writer.toString();
    }
}