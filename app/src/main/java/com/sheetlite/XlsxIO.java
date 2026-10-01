package com.sheetlite;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;
import javax.xml.parsers.DocumentBuilderFactory;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

/** Compact OOXML reader/writer for standard worksheet cells. */
final class XlsxIO {
  static Workbook read(InputStream input) throws Exception {
    Map<String, byte[]> files = new HashMap<>();
    int total = 0;
    ZipInputStream zip = new ZipInputStream(input);
    ZipEntry entry;
    byte[] buffer = new byte[8192];
    while ((entry = zip.getNextEntry()) != null) {
      if (entry.isDirectory()) continue;
      ByteArrayOutputStream out = new ByteArrayOutputStream();
      int n;
      while ((n = zip.read(buffer)) != -1) {
        out.write(buffer, 0, n);
        total += n;
        if (total > 64 * 1024 * 1024 || out.size() > 24 * 1024 * 1024)
          throw new IOException("Workbook too large");
      }
      files.put(entry.getName(), out.toByteArray());
    }
    if (!files.containsKey("xl/workbook.xml")) throw new IOException("Invalid XLSX workbook");
    List<String> shared = new ArrayList<>();
    if (files.containsKey("xl/sharedStrings.xml")) {
      Document strings = xml(files.get("xl/sharedStrings.xml"));
      NodeList items = strings.getElementsByTagName("si");
      for (int i = 0; i < items.getLength(); i++) {
        Element si = (Element) items.item(i);
        StringBuilder text = new StringBuilder();
        NodeList runs = si.getElementsByTagName("t");
        for (int j = 0; j < runs.getLength(); j++) text.append(runs.item(j).getTextContent());
        shared.add(text.toString());
      }
    }
    List<Workbook.Cell> styles = readStyles(files.get("xl/styles.xml"));
    Map<String, String> targets = new HashMap<>();
    if (files.containsKey("xl/_rels/workbook.xml.rels")) {
      NodeList rels =
          xml(files.get("xl/_rels/workbook.xml.rels")).getElementsByTagName("Relationship");
      for (int i = 0; i < rels.getLength(); i++) {
        Element e = (Element) rels.item(i);
        String target = e.getAttribute("Target");
        if (target.startsWith("/")) target = target.substring(1);
        else if (!target.startsWith("xl/")) target = "xl/" + target;
        targets.put(e.getAttribute("Id"), target);
      }
    }
    Workbook book = new Workbook();
    book.sheets.clear();
    NodeList sheets = xml(files.get("xl/workbook.xml")).getElementsByTagName("sheet");
    for (int index = 0; index < sheets.getLength(); index++) {
      Element el = (Element) sheets.item(index);
      String name = el.getAttribute("name");
      Workbook.Sheet sheet = new Workbook.Sheet(name.isEmpty() ? "Sheet" + (index + 1) : name);
      book.sheets.add(sheet);
      String path = targets.get(el.getAttribute("r:id"));
      if (path == null) path = "xl/worksheets/sheet" + (index + 1) + ".xml";
      byte[] source = files.get(path);
      if (source == null) continue;
      NodeList cells = xml(source).getElementsByTagName("c");
      for (int i = 0; i < cells.getLength(); i++) {
        Element cell = (Element) cells.item(i);
        try {
          int[] a = Workbook.address(cell.getAttribute("r"));
          String value = "";
          NodeList formula = cell.getElementsByTagName("f");
          if (formula.getLength() > 0) value = "=" + formula.item(0).getTextContent();
          else {
            String type = cell.getAttribute("t");
            NodeList v = cell.getElementsByTagName("v");
            if ("inlineStr".equals(type)) {
              NodeList texts = cell.getElementsByTagName("t");
              StringBuilder b = new StringBuilder();
              for (int j = 0; j < texts.getLength(); j++) b.append(texts.item(j).getTextContent());
              value = b.toString();
            } else if (v.getLength() > 0) {
              value = v.item(0).getTextContent();
              if ("s".equals(type)) {
                int id = Integer.parseInt(value);
                value = id >= 0 && id < shared.size() ? shared.get(id) : "";
              } else if ("b".equals(type)) value = "1".equals(value) ? "TRUE" : "FALSE";
            }
          }
          sheet.set(a[0], a[1], value);
          if (cell.hasAttribute("s")) {
            int style = Integer.parseInt(cell.getAttribute("s"));
            if (style >= 0 && style < styles.size()) {
              Workbook.Cell target = sheet.ensure(a[0], a[1]);
              target.bold = styles.get(style).bold;
              target.fill = styles.get(style).fill;
            }
          }
        } catch (RuntimeException ignored) {
          /* Skip unsupported cell data. */
        }
      }
    }
    if (book.sheets.isEmpty()) book.sheets.add(new Workbook.Sheet("Sheet1"));
    return book;
  }

  static void write(Workbook book, OutputStream output) throws IOException {
    ZipOutputStream zip = new ZipOutputStream(output);
    StringBuilder types =
        new StringBuilder(
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Types"
                + " xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default"
                + " Extension=\"rels\""
                + " ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default"
                + " Extension=\"xml\" ContentType=\"application/xml\"/><Override"
                + " PartName=\"/xl/workbook.xml\""
                + " ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override"
                + " PartName=\"/xl/styles.xml\""
                + " ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
    StringBuilder workbook =
        new StringBuilder(
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?><workbook"
                + " xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
                + " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets>");
    StringBuilder rels =
        new StringBuilder(
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships"
                + " xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
    for (int i = 0; i < book.sheets.size(); i++) {
      int id = i + 1;
      types
          .append("<Override PartName=\"/xl/worksheets/sheet")
          .append(id)
          .append(
              ".xml\""
                  + " ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
      workbook
          .append("<sheet name=\"")
          .append(escape(book.sheets.get(i).name))
          .append("\" sheetId=\"")
          .append(id)
          .append("\" r:id=\"rId")
          .append(id)
          .append("\"/>");
      rels.append("<Relationship Id=\"rId")
          .append(id)
          .append(
              "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\""
                  + " Target=\"worksheets/sheet")
          .append(id)
          .append(".xml\"/>");
      writeEntry(zip, "xl/worksheets/sheet" + id + ".xml", sheetXml(book.sheets.get(i)));
    }
    types.append("</Types>");
    workbook.append("</sheets></workbook>");
    rels.append(
        "<Relationship Id=\"rIdStyles\""
            + " Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\""
            + " Target=\"styles.xml\"/></Relationships>");
    writeEntry(zip, "[Content_Types].xml", types.toString());
    writeEntry(
        zip,
        "_rels/.rels",
        "<?xml version=\"1.0\"?><Relationships"
            + " xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship"
            + " Id=\"rId1\""
            + " Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\""
            + " Target=\"xl/workbook.xml\"/></Relationships>");
    writeEntry(zip, "xl/workbook.xml", workbook.toString());
    writeEntry(zip, "xl/_rels/workbook.xml.rels", rels.toString());
    writeEntry(zip, "xl/styles.xml", STYLES);
    zip.finish();
    zip.flush();
  }

  private static String sheetXml(Workbook.Sheet sheet) {
    StringBuilder b =
        new StringBuilder(
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet"
                + " xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>");
    int lastRow = sheet.lastRow(), lastCol = sheet.lastCol();
    for (int row = 0; row < lastRow; row++) {
      boolean opened = false;
      for (int col = 0; col < lastCol; col++) {
        String value = sheet.get(row, col);
        Workbook.Cell cell = sheet.cell(row, col);
        if (value.isEmpty() && cell == null) continue;
        if (!opened) {
          b.append("<row r=\"").append(row + 1).append("\">");
          opened = true;
        }
        b.append("<c r=\"").append(Workbook.columnName(col)).append(row + 1).append("\"");
        if (cell != null && (cell.bold || cell.fill != 0xFFFFFFFF))
          b.append(" s=\"")
              .append((cell.bold ? 1 : 0) + (cell.fill != 0xFFFFFFFF ? 2 : 0))
              .append("\"");
        if (value.startsWith("=")) {
          b.append("><f>").append(escape(value.substring(1))).append("</f><v>");
          String calculated = sheet.display(row, col);
          if (!calculated.startsWith("#")) b.append(escape(calculated));
          b.append("</v></c>");
        } else if (value.matches("-?(?:[0-9]+(?:\\.[0-9]*)?|\\.[0-9]+)"))
          b.append("><v>").append(value).append("</v></c>");
        else b.append(" t=\"inlineStr\"><is><t>").append(escape(value)).append("</t></is></c>");
      }
      if (opened) b.append("</row>");
    }
    return b.append("</sheetData></worksheet>").toString();
  }

  private static List<Workbook.Cell> readStyles(byte[] source) throws Exception {
    List<Workbook.Cell> result = new ArrayList<>();
    if (source == null) return result;
    Document doc = xml(source);
    List<Boolean> bolds = new ArrayList<>();
    NodeList fonts = doc.getElementsByTagName("fonts");
    if (fonts.getLength() > 0) {
      NodeList children = ((Element) fonts.item(0)).getElementsByTagName("font");
      for (int i = 0; i < children.getLength(); i++)
        bolds.add(((Element) children.item(i)).getElementsByTagName("b").getLength() > 0);
    }
    List<Integer> fills = new ArrayList<>();
    NodeList fillGroups = doc.getElementsByTagName("fills");
    if (fillGroups.getLength() > 0) {
      NodeList children = ((Element) fillGroups.item(0)).getElementsByTagName("fill");
      for (int i = 0; i < children.getLength(); i++) {
        int color = 0xFFFFFFFF;
        NodeList colors = ((Element) children.item(i)).getElementsByTagName("fgColor");
        if (colors.getLength() > 0) {
          String rgb = ((Element) colors.item(0)).getAttribute("rgb");
          try {
            color = (int) Long.parseLong(rgb, 16) | 0xFF000000;
          } catch (Exception ignored) {
          }
        }
        fills.add(color);
      }
    }
    NodeList groups = doc.getElementsByTagName("cellXfs");
    if (groups.getLength() > 0) {
      NodeList xfs = ((Element) groups.item(0)).getElementsByTagName("xf");
      for (int i = 0; i < xfs.getLength(); i++) {
        Element xf = (Element) xfs.item(i);
        Workbook.Cell c = new Workbook.Cell();
        try {
          int font = Integer.parseInt(xf.getAttribute("fontId")),
              fill = Integer.parseInt(xf.getAttribute("fillId"));
          c.bold = font >= 0 && font < bolds.size() && bolds.get(font);
          c.fill = fill >= 0 && fill < fills.size() ? fills.get(fill) : 0xFFFFFFFF;
        } catch (Exception ignored) {
        }
        result.add(c);
      }
    }
    return result;
  }

  private static final String STYLES =
      "<?xml version=\"1.0\" encoding=\"UTF-8\"?><styleSheet"
          + " xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><fonts"
          + " count=\"2\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font><font><b/><sz"
          + " val=\"11\"/><name val=\"Calibri\"/></font></fonts><fills"
          + " count=\"3\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill"
          + " patternType=\"gray125\"/></fill><fill><patternFill patternType=\"solid\"><fgColor"
          + " rgb=\"FFFFF2B3\"/><bgColor indexed=\"64\"/></patternFill></fill></fills><borders"
          + " count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs"
          + " count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\""
          + " borderId=\"0\"/></cellStyleXfs><cellXfs count=\"4\"><xf numFmtId=\"0\" fontId=\"0\""
          + " fillId=\"0\" borderId=\"0\" xfId=\"0\"/><xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\""
          + " borderId=\"0\" xfId=\"0\"/><xf numFmtId=\"0\" fontId=\"0\" fillId=\"2\""
          + " borderId=\"0\" xfId=\"0\"/><xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\""
          + " borderId=\"0\" xfId=\"0\"/></cellXfs><cellStyles count=\"1\"><cellStyle"
          + " name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles></styleSheet>";

  private static Document xml(byte[] data) throws Exception {
    byte[] forbidden = "<!DOCTYPE".getBytes(StandardCharsets.US_ASCII);
    for (int i = 0; i <= data.length - forbidden.length; i++) {
      int j = 0;
      while (j < forbidden.length
          && Character.toUpperCase((char) (data[i + j] & 0xFF)) == forbidden[j]) j++;
      if (j == forbidden.length) throw new IOException("DOCTYPE is not supported");
    }
    DocumentBuilderFactory f = DocumentBuilderFactory.newInstance();
    try {
      f.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
    } catch (Exception ignored) {
    }
    try {
      f.setFeature("http://xml.org/sax/features/external-general-entities", false);
    } catch (Exception ignored) {
    }
    try {
      f.setFeature("http://xml.org/sax/features/external-parameter-entities", false);
    } catch (Exception ignored) {
    }
    f.setExpandEntityReferences(false);
    return f.newDocumentBuilder().parse(new ByteArrayInputStream(data));
  }

  private static void writeEntry(ZipOutputStream zip, String name, String value)
      throws IOException {
    zip.putNextEntry(new ZipEntry(name));
    zip.write(value.getBytes(StandardCharsets.UTF_8));
    zip.closeEntry();
  }

  private static String escape(String value) {
    return value
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace("\"", "&quot;")
        .replace("'", "&apos;");
  }
}
