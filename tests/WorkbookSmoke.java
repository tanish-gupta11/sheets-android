package com.sheetlite;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;

public final class WorkbookSmoke {
  private static void check(boolean condition, String message) {
    if (!condition) throw new AssertionError(message);
  }

  public static void main(String[] args) throws Exception {
    Workbook book = new Workbook();
    Workbook.Sheet first = book.current();
    first.set(0, 0, "10");
    first.set(1, 0, "20");
    first.set(0, 1, "=SUM(A1:A2)");
    first.set(1, 1, "=A1*2+5");
    first.set(2, 1, "a,b\nquoted \"text\"");
    first.ensure(0, 0).bold = true;
    first.ensure(0, 0).fill = 0xFFFFF2B3;
    check("30".equals(first.display(0, 1)), "Range SUM");
    check("25".equals(first.display(1, 1)), "Arithmetic");
    first.set(3, 0, "=A4");
    check("#CYCLE!".equals(first.display(3, 0)), "Circular reference");
    book.addSheet();
    book.current().name = "Budget";
    book.current().set(0, 0, "Second sheet");
    ByteArrayOutputStream bytes = new ByteArrayOutputStream();
    XlsxIO.write(book, bytes);
    Workbook restored = XlsxIO.read(new ByteArrayInputStream(bytes.toByteArray()));
    check(restored.sheets.size() == 2, "Sheet count");
    check("Budget".equals(restored.sheets.get(1).name), "Sheet name");
    check("=SUM(A1:A2)".equals(restored.sheets.get(0).get(0, 1)), "Formula round trip");
    check(restored.sheets.get(0).cell(0, 0).bold, "Bold round trip");
    check(restored.sheets.get(0).cell(0, 0).fill == 0xFFFFF2B3, "Fill round trip");
    ByteArrayOutputStream csv = new ByteArrayOutputStream();
    CsvIO.write(first, csv);
    Workbook fromCsv = CsvIO.read(new ByteArrayInputStream(csv.toByteArray()));
    check("a,b\nquoted \"text\"".equals(fromCsv.current().get(2, 1)), "Quoted CSV round trip");
    if (args.length > 0)
      try (FileOutputStream file = new FileOutputStream(args[0])) {
        file.write(bytes.toByteArray());
      }
    if (args.length > 1)
      try (FileInputStream file = new FileInputStream(args[1])) {
        Workbook external = XlsxIO.read(file);
        check("Input".equals(external.current().name), "External sheet name");
        check("hello".equals(external.current().get(0, 0)), "External string");
        check("=B1*2".equals(external.current().get(0, 2)), "External formula");
      }
    ByteArrayOutputStream legacy = new ByteArrayOutputStream();
    XlsIO.write(book, legacy);
    Workbook legacyRestored = XlsIO.read(new ByteArrayInputStream(legacy.toByteArray()));
    check(legacyRestored.sheets.size() == 2, "XLS sheet count");
    check("Budget".equals(legacyRestored.sheets.get(1).name), "XLS sheet name");
    check("10".equals(legacyRestored.sheets.get(0).get(0, 0)), "XLS number");
    check("=SUM(A1:A2)".equals(legacyRestored.sheets.get(0).get(0, 1)), "XLS formula");
    if (args.length > 2)
      try (FileOutputStream file = new FileOutputStream(args[2])) {
        file.write(legacy.toByteArray());
      }
    if (args.length > 3)
      try (FileInputStream file = new FileInputStream(args[3])) {
        Workbook external = XlsIO.read(file);
        check("Legacy".equals(external.current().name), "External XLS sheet");
        check("hello".equals(external.current().get(0, 0)), "External XLS string");
        check("=B1*2".equals(external.current().get(0, 2)), "External XLS formula");
      }
    System.out.println(
        "Workbook smoke checks passed; XLSX bytes="
            + bytes.size()
            + ", XLS bytes="
            + legacy.size());
  }
}
