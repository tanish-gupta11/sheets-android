package com.sheetlite;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.nio.charset.StandardCharsets;

final class CsvIO {
  static Workbook read(InputStream input) throws IOException {
    Workbook book = new Workbook();
    Workbook.Sheet sheet = book.current();
    BufferedReader reader =
        new BufferedReader(new InputStreamReader(input, StandardCharsets.UTF_8));
    StringBuilder field = new StringBuilder();
    int row = 0, col = 0;
    boolean quoted = false;
    int next;
    while ((next = reader.read()) != -1) {
      char ch = (char) next;
      if (quoted) {
        if (ch == '"') {
          reader.mark(1);
          int following = reader.read();
          if (following == '"') field.append('"');
          else {
            quoted = false;
            if (following != -1) reader.reset();
          }
        } else field.append(ch);
      } else if (ch == '"' && field.length() == 0) quoted = true;
      else if (ch == ',') {
        sheet.set(row, col++, field.toString());
        field.setLength(0);
      } else if (ch == '\n') {
        sheet.set(row, col, field.toString());
        field.setLength(0);
        row++;
        col = 0;
        if (row >= Workbook.MAX_ROWS) break;
      } else if (ch != '\r') field.append(ch);
    }
    if (quoted) throw new IOException("Unclosed CSV quote");
    if (field.length() > 0 || col > 0) sheet.set(row, col, field.toString());
    return book;
  }

  static void write(Workbook.Sheet sheet, OutputStream output) throws IOException {
    BufferedWriter writer =
        new BufferedWriter(new OutputStreamWriter(output, StandardCharsets.UTF_8));
    int lastRow = sheet.lastRow(), lastCol = sheet.lastCol();
    for (int r = 0; r < lastRow; r++) {
      for (int c = 0; c < lastCol; c++) {
        if (c > 0) writer.write(',');
        String v = sheet.get(r, c);
        if (v.indexOf(',') >= 0
            || v.indexOf('"') >= 0
            || v.indexOf('\n') >= 0
            || v.indexOf('\r') >= 0) {
          writer.write('"');
          writer.write(v.replace("\"", "\"\""));
          writer.write('"');
        } else writer.write(v);
      }
      writer.write("\r\n");
    }
    writer.flush();
  }
}
