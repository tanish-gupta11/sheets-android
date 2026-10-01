package com.sheetlite;

import static org.junit.Assert.*;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import org.junit.Test;

public final class WorkbookTest {
  @Test
  public void formulasAndCircularReferences() {
    Workbook.Sheet sheet = new Workbook.Sheet("Test");
    sheet.set(0, 0, "10");
    sheet.set(1, 0, "20");
    sheet.set(0, 1, "=SUM(A1:A2)");
    assertEquals("30", sheet.display(0, 1));
    sheet.set(2, 0, "=A3");
    assertEquals("#CYCLE!", sheet.display(2, 0));
  }

  @Test
  public void csvQuotedFields() throws Exception {
    Workbook.Sheet sheet = new Workbook.Sheet("Test");
    sheet.set(0, 0, "a,b\n\"quoted\"");
    ByteArrayOutputStream out = new ByteArrayOutputStream();
    CsvIO.write(sheet, out);
    assertEquals(
        sheet.get(0, 0),
        CsvIO.read(new ByteArrayInputStream(out.toByteArray())).current().get(0, 0));
  }

  @Test
  public void structuralEditsUpdateCellReferences() {
    Workbook.Sheet sheet = new Workbook.Sheet("Test");
    sheet.set(0, 0, "5");
    sheet.set(0, 1, "=A1*2");
    sheet.insertRow(0);
    assertEquals("=A2*2", sheet.get(1, 1));
    assertEquals("10", sheet.display(1, 1));
    sheet.insertCol(0);
    assertEquals("=B2*2", sheet.get(1, 2));
    sheet.deleteCol(1);
    assertEquals("=#REF!*2", sheet.get(1, 1));
    assertEquals("#REF!", sheet.display(1, 1));
  }

  @Test
  public void xlsxRoundTrip() throws Exception {
    Workbook book = sample();
    ByteArrayOutputStream out = new ByteArrayOutputStream();
    XlsxIO.write(book, out);
    Workbook copy = XlsxIO.read(new ByteArrayInputStream(out.toByteArray()));
    assertEquals(2, copy.sheets.size());
    assertEquals("Budget", copy.sheets.get(1).name);
    assertEquals("=A1*2", copy.sheets.get(0).get(0, 1));
    assertTrue(copy.sheets.get(0).cell(0, 0).bold);
  }

  @Test
  public void xlsRoundTrip() throws Exception {
    Workbook book = sample();
    ByteArrayOutputStream out = new ByteArrayOutputStream();
    XlsIO.write(book, out);
    Workbook copy = XlsIO.read(new ByteArrayInputStream(out.toByteArray()));
    assertEquals(2, copy.sheets.size());
    assertEquals("Budget", copy.sheets.get(1).name);
    assertEquals("=A1*2", copy.sheets.get(0).get(0, 1));
    assertTrue(copy.sheets.get(0).cell(0, 0).bold);
  }

  private Workbook sample() {
    Workbook book = new Workbook();
    book.current().set(0, 0, "7");
    book.current().ensure(0, 0).bold = true;
    book.current().set(0, 1, "=A1*2");
    book.addSheet();
    book.current().name = "Budget";
    return book;
  }
}
