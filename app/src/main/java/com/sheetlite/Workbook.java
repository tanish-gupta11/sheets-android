package com.sheetlite;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

final class Workbook implements Serializable {
  private static final long serialVersionUID = 2L;
  static final int MAX_ROWS = 10000, MAX_COLS = 256;
  final List<Sheet> sheets = new ArrayList<>();
  int active = 0;

  Workbook() {
    sheets.add(new Sheet("Sheet1"));
  }

  Sheet current() {
    return sheets.get(active);
  }

  void addSheet() {
    sheets.add(new Sheet("Sheet" + (sheets.size() + 1)));
    active = sheets.size() - 1;
  }

  static String columnName(int col) {
    StringBuilder out = new StringBuilder();
    do {
      out.append((char) ('A' + col % 26));
      col = col / 26 - 1;
    } while (col >= 0);
    return out.reverse().toString();
  }

  static int[] address(String address) {
    int i = 0, col = 0;
    while (i < address.length() && Character.isLetter(address.charAt(i))) {
      col = col * 26 + (Character.toUpperCase(address.charAt(i)) - 'A' + 1);
      i++;
    }
    if (i == 0 || i == address.length() || col < 1)
      throw new IllegalArgumentException("Invalid address");
    int row = Integer.parseInt(address.substring(i));
    if (row < 1 || row > MAX_ROWS || col > MAX_COLS)
      throw new IllegalArgumentException("Out of range");
    return new int[] {row - 1, col - 1};
  }

  static final class Cell implements Serializable {
    private static final long serialVersionUID = 1L;
    String value = "";
    boolean bold;
    int fill = 0xFFFFFFFF;

    Cell copy() {
      Cell c = new Cell();
      c.value = value;
      c.bold = bold;
      c.fill = fill;
      return c;
    }
  }

  static final class Sheet implements Serializable {
    private static final long serialVersionUID = 1L;
    private static final Pattern CELL_REFERENCE =
        Pattern.compile(
            "(?i)(?<![A-Z0-9_!])(\\$?)([A-Z]{1,3})(\\$?)([1-9][0-9]{0,4})(?![A-Z0-9_])");
    String name;
    final Map<Long, Cell> cells = new HashMap<>();
    int rows = 1000, cols = 52;

    Sheet(String name) {
      this.name = name;
    }

    static long key(int r, int c) {
      return ((long) r << 32) | (c & 0xffffffffL);
    }

    Cell cell(int r, int c) {
      return cells.get(key(r, c));
    }

    Cell ensure(int r, int c) {
      long k = key(r, c);
      Cell v = cells.get(k);
      if (v == null) {
        v = new Cell();
        cells.put(k, v);
      }
      return v;
    }

    String get(int r, int c) {
      Cell v = cell(r, c);
      return v == null ? "" : v.value;
    }

    void set(int r, int c, String value) {
      if (r < 0 || r >= MAX_ROWS || c < 0 || c >= MAX_COLS) return;
      if (value == null || value.isEmpty()) {
        Cell existing = cell(r, c);
        if (existing != null) {
          existing.value = "";
          if (!existing.bold && existing.fill == 0xFFFFFFFF) cells.remove(key(r, c));
        }
      } else ensure(r, c).value = value;
      rows = Math.max(rows, r + 1);
      cols = Math.max(cols, c + 1);
    }

    String display(int r, int c) {
      String v = get(r, c);
      return v.startsWith("=") ? FormulaEngine.evaluate(v, this) : v;
    }

    int lastRow() {
      int n = 0;
      for (long key : cells.keySet()) n = Math.max(n, (int) (key >>> 32) + 1);
      return n;
    }

    int lastCol() {
      int n = 0;
      for (long key : cells.keySet()) n = Math.max(n, (int) (long) key + 1);
      return n;
    }

    void insertRow(int row) {
      shift(row, true, true);
    }

    void deleteRow(int row) {
      shift(row, true, false);
    }

    void insertCol(int col) {
      shift(col, false, true);
    }

    void deleteCol(int col) {
      shift(col, false, false);
    }

    private void shift(int index, boolean rowAxis, boolean insert) {
      Map<Long, Cell> next = new HashMap<>();
      for (Map.Entry<Long, Cell> e : cells.entrySet()) {
        int r = (int) (e.getKey() >>> 32), c = (int) (long) e.getKey();
        int pos = rowAxis ? r : c;
        if (!insert && pos == index) continue;
        if (insert && pos >= index) pos++;
        if (!insert && pos > index) pos--;
        if (rowAxis) r = pos;
        else c = pos;
        if (r >= 0 && r < MAX_ROWS && c >= 0 && c < MAX_COLS) {
          Cell cell = e.getValue();
          if (cell.value.startsWith("=")) {
            cell.value = adjustReferences(cell.value, index, rowAxis, insert);
          }
          next.put(key(r, c), cell);
        }
      }
      cells.clear();
      cells.putAll(next);
      if (rowAxis) rows = insert ? Math.min(MAX_ROWS, rows + 1) : Math.max(1, rows - 1);
      else cols = insert ? Math.min(MAX_COLS, cols + 1) : Math.max(1, cols - 1);
    }

    private static String adjustReferences(
        String formula, int index, boolean rowAxis, boolean insert) {
      Matcher matcher = CELL_REFERENCE.matcher(formula);
      StringBuffer result = new StringBuffer();
      while (matcher.find()) {
        String original = matcher.group();
        String replacement = original;
        try {
          int[] address = Workbook.address(matcher.group(2) + matcher.group(4));
          int position = rowAxis ? address[0] : address[1];
          if (!insert && position == index) replacement = "#REF!";
          else if (position >= index) {
            int next = position + (insert ? 1 : -1);
            if (next < 0 || next >= (rowAxis ? MAX_ROWS : MAX_COLS)) replacement = "#REF!";
            else if (rowAxis)
              replacement = matcher.group(1) + matcher.group(2) + matcher.group(3) + (next + 1);
            else
              replacement =
                  matcher.group(1)
                      + Workbook.columnName(next)
                      + matcher.group(3)
                      + matcher.group(4);
          }
        } catch (IllegalArgumentException ignored) {
          // Function names and names outside the supported grid remain unchanged.
        }
        matcher.appendReplacement(result, Matcher.quoteReplacement(replacement));
      }
      matcher.appendTail(result);
      return result.toString();
    }
  }
}
