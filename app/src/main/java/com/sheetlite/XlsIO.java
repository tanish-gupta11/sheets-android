package com.sheetlite;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.Map;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.WorkbookUtil;

/** Legacy Excel 97–2003 BIFF8 files, backed by Apache POI HSSF. */
final class XlsIO {
  static Workbook read(InputStream stream) throws IOException {
    ByteArrayOutputStream bounded = new ByteArrayOutputStream();
    byte[] chunk = new byte[8192];
    int count;
    while ((count = stream.read(chunk)) != -1) {
      bounded.write(chunk, 0, count);
      if (bounded.size() > 32 * 1024 * 1024) throw new IOException("XLS file exceeds 32 MB limit");
    }
    try (HSSFWorkbook source = new HSSFWorkbook(new ByteArrayInputStream(bounded.toByteArray()))) {
      Workbook result = new Workbook();
      result.sheets.clear();
      for (int s = 0; s < source.getNumberOfSheets(); s++) {
        HSSFSheet from = source.getSheetAt(s);
        Workbook.Sheet to = new Workbook.Sheet(from.getSheetName());
        result.sheets.add(to);
        int lastRow = Math.min(Workbook.MAX_ROWS - 1, from.getLastRowNum());
        for (int r = 0; r <= lastRow; r++) {
          HSSFRow row = from.getRow(r);
          if (row == null) continue;
          int lastCol = Math.min(Workbook.MAX_COLS, row.getLastCellNum());
          for (int c = 0; c < lastCol; c++) {
            HSSFCell cell = row.getCell(c);
            if (cell == null) continue;
            String value = value(cell);
            if (!value.isEmpty()) to.set(r, c, value);
            HSSFCellStyle style = cell.getCellStyle();
            if (style != null) {
              boolean bold = source.getFontAt(style.getFontIndexAsInt()).getBold();
              boolean filled = style.getFillPattern() == FillPatternType.SOLID_FOREGROUND;
              if (bold || filled) {
                Workbook.Cell target = to.ensure(r, c);
                target.bold = bold;
                if (filled) target.fill = 0xFFFFF2B3;
              }
            }
          }
        }
      }
      if (result.sheets.isEmpty()) result.sheets.add(new Workbook.Sheet("Sheet1"));
      return result;
    } catch (RuntimeException error) {
      throw new IOException("Unsupported or damaged XLS file", error);
    }
  }

  private static String value(HSSFCell cell) {
    CellType type = cell.getCellType();
    switch (type) {
      case STRING:
        return cell.getStringCellValue();
      case NUMERIC:
        return BigDecimal.valueOf(cell.getNumericCellValue()).stripTrailingZeros().toPlainString();
      case BOOLEAN:
        return cell.getBooleanCellValue() ? "TRUE" : "FALSE";
      case FORMULA:
        return "=" + cell.getCellFormula();
      case ERROR:
        return "#ERROR!";
      default:
        return "";
    }
  }

  static void write(Workbook source, OutputStream stream) throws IOException {
    try (HSSFWorkbook output = new HSSFWorkbook()) {
      HSSFFont bold = output.createFont();
      bold.setBold(true);
      HSSFCellStyle normal = output.createCellStyle(),
          boldStyle = output.createCellStyle(),
          yellow = output.createCellStyle(),
          boldYellow = output.createCellStyle();
      boldStyle.setFont(bold);
      yellow.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
      yellow.setFillPattern(FillPatternType.SOLID_FOREGROUND);
      boldYellow.cloneStyleFrom(yellow);
      boldYellow.setFont(bold);
      for (int s = 0; s < source.sheets.size(); s++) {
        Workbook.Sheet sheet = source.sheets.get(s);
        String base = WorkbookUtil.createSafeSheetName(sheet.name);
        if (base.isEmpty()) base = "Sheet" + (s + 1);
        if (base.length() > 31) base = base.substring(0, 31);
        String safe = base;
        int suffix = 2;
        while (output.getSheet(safe) != null) {
          String tail = "_" + suffix++;
          safe = base.substring(0, Math.min(base.length(), 31 - tail.length())) + tail;
        }
        HSSFSheet target = output.createSheet(safe);
        for (Map.Entry<Long, Workbook.Cell> entry : sheet.cells.entrySet()) {
          int r = (int) (entry.getKey() >>> 32), c = (int) (long) entry.getKey();
          if (r >= Workbook.MAX_ROWS || c >= Workbook.MAX_COLS) continue;
          HSSFRow row = target.getRow(r);
          if (row == null) row = target.createRow(r);
          HSSFCell cell = row.createCell(c);
          Workbook.Cell from = entry.getValue();
          cell.setCellStyle(
              from.bold
                  ? (from.fill == 0xFFFFFFFF ? boldStyle : boldYellow)
                  : (from.fill == 0xFFFFFFFF ? normal : yellow));
          String v = from.value;
          if (v.startsWith("=")) {
            try {
              cell.setCellFormula(v.substring(1));
            } catch (Exception unsupported) {
              cell.setCellValue(v);
            }
          } else if (v.matches("-?(?:[0-9]+(?:\\.[0-9]*)?|\\.[0-9]+)")) {
            try {
              cell.setCellValue(Double.parseDouble(v));
            } catch (Exception error) {
              cell.setCellValue(v);
            }
          } else cell.setCellValue(v);
        }
      }
      output.write(stream);
      stream.flush();
    } catch (RuntimeException error) {
      throw new IOException("Could not write XLS file", error);
    }
  }
}
