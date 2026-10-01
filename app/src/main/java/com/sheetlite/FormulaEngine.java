package com.sheetlite;

import java.util.HashSet;
import java.util.Locale;
import java.util.Set;

final class FormulaEngine {
  static String evaluate(String raw, Workbook.Sheet sheet) {
    if (raw.contains("#REF!")) return "#REF!";
    try {
      double value = new Parser(raw.substring(1), sheet, new HashSet<Long>()).parse();
      if (!Double.isFinite(value)) return "#DIV/0!";
      return value == Math.rint(value) && Math.abs(value) < 9e15
          ? Long.toString((long) value)
          : String.format(Locale.US, "%.8f", value).replaceAll("0+$", "").replaceAll("\\.$", "");
    } catch (Cycle e) {
      return "#CYCLE!";
    } catch (Exception e) {
      return "#ERROR!";
    }
  }

  private static final class Cycle extends RuntimeException {}

  private static final class Parser {
    final String s;
    final Workbook.Sheet sheet;
    final Set<Long> visiting;
    int p;

    Parser(String s, Workbook.Sheet sheet, Set<Long> visiting) {
      this.s = s.toUpperCase(Locale.US).replace("$", "");
      this.sheet = sheet;
      this.visiting = visiting;
    }

    double parse() {
      double v = expr();
      space();
      if (p != s.length()) throw new IllegalArgumentException();
      return v;
    }

    void space() {
      while (p < s.length() && Character.isWhitespace(s.charAt(p))) p++;
    }

    boolean eat(char ch) {
      space();
      if (p < s.length() && s.charAt(p) == ch) {
        p++;
        return true;
      }
      return false;
    }

    double expr() {
      double v = term();
      for (; ; ) {
        if (eat('+')) v += term();
        else if (eat('-')) v -= term();
        else return v;
      }
    }

    double term() {
      double v = power();
      for (; ; ) {
        if (eat('*')) v *= power();
        else if (eat('/')) v /= power();
        else return v;
      }
    }

    double power() {
      double v = factor();
      if (eat('^')) v = Math.pow(v, power());
      return v;
    }

    double factor() {
      space();
      if (eat('+')) return factor();
      if (eat('-')) return -factor();
      if (eat('(')) {
        double v = expr();
        if (!eat(')')) throw new IllegalArgumentException();
        return v;
      }
      if (p >= s.length()) throw new IllegalArgumentException();
      if (Character.isDigit(s.charAt(p)) || s.charAt(p) == '.') {
        int start = p;
        while (p < s.length() && (Character.isDigit(s.charAt(p)) || s.charAt(p) == '.')) p++;
        return Double.parseDouble(s.substring(start, p));
      }
      if (Character.isLetter(s.charAt(p))) {
        int start = p;
        while (p < s.length() && Character.isLetter(s.charAt(p))) p++;
        String word = s.substring(start, p);
        if (eat('(')) return function(word);
        int rowStart = p;
        while (p < s.length() && Character.isDigit(s.charAt(p))) p++;
        if (rowStart == p) throw new IllegalArgumentException();
        int[] a = Workbook.address(s.substring(start, p));
        return cell(a[0], a[1]);
      }
      throw new IllegalArgumentException();
    }

    double cell(int row, int col) {
      long key = Workbook.Sheet.key(row, col);
      if (!visiting.add(key)) throw new Cycle();
      try {
        String v = sheet.get(row, col);
        if (v.isEmpty()) return 0;
        if (v.startsWith("=")) return new Parser(v.substring(1), sheet, visiting).parse();
        try {
          return Double.parseDouble(v);
        } catch (NumberFormatException e) {
          return 0;
        }
      } finally {
        visiting.remove(key);
      }
    }

    double function(String name) {
      double total = 0, min = Double.POSITIVE_INFINITY, max = Double.NEGATIVE_INFINITY;
      int count = 0;
      boolean all = true, any = false;
      do {
        space();
        int begin = p;
        double value = expr();
        space();
        if (eat(':')) {
          int endBegin = p;
          while (p < s.length() && Character.isLetterOrDigit(s.charAt(p))) p++;
          int[] a = Workbook.address(s.substring(begin, endBegin - 1).trim());
          int[] b = Workbook.address(s.substring(endBegin, p).trim());
          if ((long) (Math.abs(a[0] - b[0]) + 1) * (Math.abs(a[1] - b[1]) + 1) > 100000)
            throw new IllegalArgumentException();
          for (int r = Math.min(a[0], b[0]); r <= Math.max(a[0], b[0]); r++)
            for (int c = Math.min(a[1], b[1]); c <= Math.max(a[1], b[1]); c++) {
              double x = cell(r, c);
              total += x;
              min = Math.min(min, x);
              max = Math.max(max, x);
              count++;
              all &= x != 0;
              any |= x != 0;
            }
        } else {
          total += value;
          min = Math.min(min, value);
          max = Math.max(max, value);
          count++;
          all &= value != 0;
          any |= value != 0;
        }
      } while (eat(','));
      if (!eat(')')) throw new IllegalArgumentException();
      switch (name) {
        case "SUM":
          return total;
        case "AVERAGE":
        case "AVG":
          return count == 0 ? 0 : total / count;
        case "MIN":
          return count == 0 ? 0 : min;
        case "MAX":
          return count == 0 ? 0 : max;
        case "COUNT":
          return count;
        case "ABS":
          return Math.abs(total);
        case "ROUND":
          return Math.rint(total);
        case "AND":
          return all ? 1 : 0;
        case "OR":
          return any ? 1 : 0;
        default:
          throw new IllegalArgumentException();
      }
    }
  }
}
