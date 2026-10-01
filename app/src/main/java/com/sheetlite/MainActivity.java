package com.sheetlite;

import android.app.Activity;
import android.app.AlertDialog;
import android.content.Context;
import android.content.Intent;
import android.graphics.Canvas;
import android.graphics.Color;
import android.graphics.Paint;
import android.graphics.Typeface;
import android.net.Uri;
import android.os.Bundle;
import android.view.Gravity;
import android.view.MotionEvent;
import android.view.View;
import android.view.WindowManager;
import android.widget.Button;
import android.widget.EditText;
import android.widget.HorizontalScrollView;
import android.widget.LinearLayout;
import android.widget.TextView;
import android.widget.Toast;
import java.io.BufferedInputStream;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;
import java.io.OutputStream;
import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public final class MainActivity extends Activity {
  private static final int OPEN = 10, SAVE_XLSX = 11, SAVE_CSV = 12, SAVE_XLS = 13;
  private Workbook book = new Workbook();
  private final ArrayDeque<byte[]> undo = new ArrayDeque<>(), redo = new ArrayDeque<>();
  private Grid grid;
  private TextView location;
  private EditText formula;
  private LinearLayout tabs;
  private int selectedRow, selectedCol;

  @Override
  public void onCreate(Bundle state) {
    super.onCreate(state);
    try (ObjectInputStream in = new ObjectInputStream(openFileInput("autosave.bin"))) {
      book = (Workbook) in.readObject();
      if (book.sheets.isEmpty()) book = new Workbook();
    } catch (Exception ignored) {
    }
    build();
  }

  @Override
  protected void onStop() {
    super.onStop();
    persist();
  }

  private void persist() {
    try (ObjectOutputStream out =
        new ObjectOutputStream(openFileOutput("autosave.bin", MODE_PRIVATE))) {
      out.writeObject(book);
    } catch (Exception ignored) {
    }
  }

  private int dp(float n) {
    return (int) (n * getResources().getDisplayMetrics().density + .5f);
  }

  private TextView text(String label, int size, int color) {
    TextView v = new TextView(this);
    v.setText(label);
    v.setTextSize(size);
    v.setTextColor(color);
    v.setGravity(Gravity.CENTER_VERTICAL);
    return v;
  }

  private Button button(String label, View.OnClickListener action) {
    Button b = new Button(this);
    b.setText(label);
    b.setAllCaps(false);
    b.setTextSize(12);
    b.setOnClickListener(action);
    return b;
  }

  private void build() {
    LinearLayout root = new LinearLayout(this);
    root.setOrientation(1);
    root.setBackgroundColor(Color.WHITE);
    root.setFitsSystemWindows(true);
    LinearLayout toolbar = new LinearLayout(this);
    toolbar.setPadding(dp(8), dp(5), dp(8), dp(3));
    toolbar.setGravity(Gravity.CENTER_VERTICAL);
    TextView title = text("▦ Sheets", 19, 0xFF173248);
    title.setTypeface(null, Typeface.BOLD);
    toolbar.addView(title, new LinearLayout.LayoutParams(0, dp(46), 1));
    toolbar.addView(button("Open", v -> open()));
    toolbar.addView(button("Save", v -> saveMenu()));
    toolbar.addView(button("Tools", v -> tools()));
    root.addView(toolbar);
    LinearLayout entry = new LinearLayout(this);
    entry.setBackgroundColor(0xFFEAF2F6);
    location = text("A1", 14, 0xFF135F81);
    location.setGravity(Gravity.CENTER);
    location.setTypeface(null, Typeface.BOLD);
    entry.addView(location, new LinearLayout.LayoutParams(dp(64), dp(45)));
    formula = new EditText(this);
    formula.setSingleLine(true);
    formula.setTextSize(15);
    formula.setHint("Cell value or formula");
    formula.setBackgroundColor(Color.TRANSPARENT);
    formula.setPadding(dp(7), 0, dp(6), 0);
    entry.addView(formula, new LinearLayout.LayoutParams(0, dp(45), 1));
    entry.addView(button("✓", v -> applyFormula()), new LinearLayout.LayoutParams(dp(46), dp(45)));
    root.addView(entry);
    grid = new Grid(this);
    root.addView(grid, new LinearLayout.LayoutParams(-1, 0, 1));
    HorizontalScrollView strip = new HorizontalScrollView(this);
    strip.setHorizontalScrollBarEnabled(false);
    tabs = new LinearLayout(this);
    tabs.setPadding(dp(6), 0, dp(6), 0);
    strip.addView(tabs);
    root.addView(strip, new LinearLayout.LayoutParams(-1, dp(48)));
    setContentView(root);
    refresh();
  }

  private void refresh() {
    if (book.active >= book.sheets.size()) book.active = 0;
    location.setText(Workbook.columnName(selectedCol) + (selectedRow + 1));
    formula.setText(book.current().get(selectedRow, selectedCol));
    tabs.removeAllViews();
    for (int i = 0; i < book.sheets.size(); i++) {
      final int index = i;
      Button tab =
          button(
              book.sheets.get(i).name,
              v -> {
                book.active = index;
                selectedRow = selectedCol = 0;
                grid.reset();
                refresh();
              });
      tab.setTextColor(i == book.active ? 0xFF126C92 : 0xFF465665);
      tabs.addView(tab);
    }
    tabs.addView(button("＋", v -> mutate(() -> book.addSheet())));
    grid.invalidate();
  }

  private void applyFormula() {
    String value = formula.getText().toString();
    mutate(() -> book.current().set(selectedRow, selectedCol, value));
  }

  private void editCell() {
    EditText input = new EditText(this);
    input.setText(book.current().get(selectedRow, selectedCol));
    input.setSelectAllOnFocus(true);
    input.setSingleLine(false);
    AlertDialog dialog =
        new AlertDialog.Builder(this)
            .setTitle(Workbook.columnName(selectedCol) + (selectedRow + 1))
            .setView(input)
            .setNegativeButton("Cancel", null)
            .setNeutralButton(
                "Clear", (d, w) -> mutate(() -> book.current().set(selectedRow, selectedCol, "")))
            .setPositiveButton(
                "Apply",
                (d, w) ->
                    mutate(
                        () ->
                            book.current()
                                .set(selectedRow, selectedCol, input.getText().toString())))
            .create();
    dialog.show();
    input.requestFocus();
    dialog.getWindow().setSoftInputMode(WindowManager.LayoutParams.SOFT_INPUT_STATE_ALWAYS_VISIBLE);
  }

  private byte[] snapshot() {
    try {
      ByteArrayOutputStream b = new ByteArrayOutputStream();
      ObjectOutputStream o = new ObjectOutputStream(b);
      o.writeObject(book);
      o.close();
      return b.toByteArray();
    } catch (Exception e) {
      return null;
    }
  }

  private void restore(byte[] data) {
    if (data == null) return;
    try {
      book = (Workbook) new ObjectInputStream(new ByteArrayInputStream(data)).readObject();
      refresh();
    } catch (Exception e) {
      toast("Could not restore change");
    }
  }

  private void mutate(Runnable action) {
    byte[] previous = snapshot();
    if (previous != null) {
      undo.push(previous);
      if (undo.size() > 20) undo.removeLast();
    }
    redo.clear();
    action.run();
    refresh();
  }

  private void undo() {
    if (undo.isEmpty()) {
      toast("Nothing to undo");
      return;
    }
    redo.push(snapshot());
    restore(undo.pop());
  }

  private void redo() {
    if (redo.isEmpty()) {
      toast("Nothing to redo");
      return;
    }
    undo.push(snapshot());
    restore(redo.pop());
  }

  private void toast(String message) {
    Toast.makeText(this, message, Toast.LENGTH_SHORT).show();
  }

  private void open() {
    Intent i = new Intent(Intent.ACTION_OPEN_DOCUMENT);
    i.addCategory(Intent.CATEGORY_OPENABLE);
    i.setType("*/*");
    i.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
    startActivityForResult(i, OPEN);
  }

  private void saveMenu() {
    new AlertDialog.Builder(this)
        .setTitle("Save workbook")
        .setItems(
            new String[] {
              "Excel workbook (.xlsx)", "Legacy Excel workbook (.xls)", "Current sheet (.csv)"
            },
            (d, which) -> {
              Intent i = new Intent(Intent.ACTION_CREATE_DOCUMENT);
              i.addCategory(Intent.CATEGORY_OPENABLE);
              i.setType(
                  which == 0
                      ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                      : which == 1 ? "application/vnd.ms-excel" : "text/csv");
              i.putExtra(
                  Intent.EXTRA_TITLE,
                  which == 0 ? "sheets.xlsx" : which == 1 ? "sheets.xls" : "sheets.csv");
              startActivityForResult(i, which == 0 ? SAVE_XLSX : which == 1 ? SAVE_XLS : SAVE_CSV);
            })
        .show();
  }

  @Override
  protected void onActivityResult(int req, int result, Intent data) {
    super.onActivityResult(req, result, data);
    if (result != RESULT_OK || data == null) return;
    Uri uri = data.getData();
    if (uri == null) return;
    try {
      if (req == OPEN) {
        try (BufferedInputStream in =
            new BufferedInputStream(getContentResolver().openInputStream(uri))) {
          in.mark(8);
          byte[] header = new byte[8];
          int read = in.read(header);
          in.reset();
          boolean xlsx = read >= 2 && header[0] == 'P' && header[1] == 'K';
          boolean xls =
              read == 8
                  && (header[0] & 0xFF) == 0xD0
                  && (header[1] & 0xFF) == 0xCF
                  && (header[2] & 0xFF) == 0x11
                  && (header[3] & 0xFF) == 0xE0;
          Workbook imported = xlsx ? XlsxIO.read(in) : xls ? XlsIO.read(in) : CsvIO.read(in);
          byte[] before = snapshot();
          if (before != null) undo.push(before);
          redo.clear();
          book = imported;
          selectedRow = selectedCol = 0;
          refresh();
          if (xlsx || xls)
            new AlertDialog.Builder(this)
                .setTitle("Workbook imported")
                .setMessage(
                    "Cells, formulas, sheet names, bold text and solid fills were imported where"
                        + " supported. Charts, pivot tables, merges, macros and advanced formatting"
                        + " are not preserved when you save a copy from this app.")
                .setPositiveButton("OK", null)
                .show();
          else toast("CSV opened");
        }
      } else if (req == SAVE_XLSX || req == SAVE_XLS || req == SAVE_CSV) {
        try (OutputStream out = getContentResolver().openOutputStream(uri)) {
          if (out == null) throw new Exception("Cannot write file");
          if (req == SAVE_XLSX) XlsxIO.write(book, out);
          else if (req == SAVE_XLS) XlsIO.write(book, out);
          else CsvIO.write(book.current(), out);
          toast("Saved");
        }
      }
    } catch (Exception e) {
      toast("File error: " + e.getMessage());
    }
  }

  private void tools() {
    String[] choices = {
      "Undo",
      "Redo",
      "Find",
      "Replace",
      "Go to cell",
      "Bold cell",
      "Yellow fill",
      "Clear fill",
      "Insert row",
      "Delete row",
      "Insert column",
      "Delete column",
      "Sort rows A–Z by column",
      "Sort rows Z–A by column",
      "Rename sheet",
      "Add sheet",
      "Delete sheet"
    };
    new AlertDialog.Builder(this)
        .setTitle("Sheet tools")
        .setItems(
            choices,
            (d, n) -> {
              switch (n) {
                case 0:
                  undo();
                  break;
                case 1:
                  redo();
                  break;
                case 2:
                  find(false);
                  break;
                case 3:
                  find(true);
                  break;
                case 4:
                  ask(
                      "Go to cell",
                      "A1",
                      s -> {
                        try {
                          int[] a = Workbook.address(s.trim());
                          selectedRow = a[0];
                          selectedCol = a[1];
                          book.current().rows = Math.max(book.current().rows, selectedRow + 1);
                          book.current().cols = Math.max(book.current().cols, selectedCol + 1);
                          grid.showCell();
                          refresh();
                        } catch (Exception e) {
                          toast("Use a cell such as B12");
                        }
                      });
                  break;
                case 5:
                  mutate(
                      () -> {
                        Workbook.Cell c = book.current().ensure(selectedRow, selectedCol);
                        c.bold = !c.bold;
                      });
                  break;
                case 6:
                  mutate(() -> book.current().ensure(selectedRow, selectedCol).fill = 0xFFFFF2B3);
                  break;
                case 7:
                  mutate(() -> book.current().ensure(selectedRow, selectedCol).fill = Color.WHITE);
                  break;
                case 8:
                  mutate(() -> book.current().insertRow(selectedRow));
                  break;
                case 9:
                  mutate(() -> book.current().deleteRow(selectedRow));
                  break;
                case 10:
                  mutate(() -> book.current().insertCol(selectedCol));
                  break;
                case 11:
                  mutate(() -> book.current().deleteCol(selectedCol));
                  break;
                case 12:
                  sort(false);
                  break;
                case 13:
                  sort(true);
                  break;
                case 14:
                  ask(
                      "Rename sheet",
                      book.current().name,
                      s ->
                          mutate(
                              () ->
                                  book.current().name =
                                      s.trim().isEmpty() ? book.current().name : s.trim()));
                  break;
                case 15:
                  mutate(() -> book.addSheet());
                  break;
                case 16:
                  if (book.sheets.size() > 1)
                    new AlertDialog.Builder(this)
                        .setMessage("Delete " + book.current().name + "?")
                        .setNegativeButton("Cancel", null)
                        .setPositiveButton(
                            "Delete",
                            (x, y) ->
                                mutate(
                                    () -> {
                                      book.sheets.remove(book.active);
                                      book.active = 0;
                                    }))
                        .show();
                  break;
              }
            })
        .show();
  }

  private interface Answer {
    void accept(String value);
  }

  private void ask(String title, String initial, Answer answer) {
    EditText input = new EditText(this);
    input.setSingleLine(true);
    input.setText(initial);
    input.selectAll();
    new AlertDialog.Builder(this)
        .setTitle(title)
        .setView(input)
        .setNegativeButton("Cancel", null)
        .setPositiveButton("OK", (d, w) -> answer.accept(input.getText().toString()))
        .show();
  }

  private void find(boolean replace) {
    ask(
        "Find text",
        "",
        query -> {
          if (query.isEmpty()) return;
          if (replace) {
            ask(
                "Replace with",
                "",
                replacement ->
                    mutate(
                        () -> {
                          Workbook.Sheet s = book.current();
                          for (Workbook.Cell c : s.cells.values())
                            if (!c.value.startsWith("="))
                              c.value = c.value.replace(query, replacement);
                        }));
          } else {
            Workbook.Sheet s = book.current();
            for (int r = selectedRow; r < s.lastRow(); r++)
              for (int c = 0; c < s.lastCol(); c++)
                if (s.get(r, c).toLowerCase().contains(query.toLowerCase())) {
                  selectedRow = r;
                  selectedCol = c;
                  grid.showCell();
                  refresh();
                  return;
                }
            toast("No match");
          }
        });
  }

  private void sort(boolean descending) {
    Workbook.Sheet sheet = book.current();
    int last = sheet.lastRow(), col = selectedCol;
    if (last < 2) return;
    mutate(
        () -> {
          List<Integer> order = new ArrayList<>();
          for (int r = 0; r < last; r++) order.add(r);
          Collections.sort(
              order,
              (a, b) -> {
                String av = sheet.display(a, col), bv = sheet.display(b, col);
                int cmp;
                try {
                  cmp = Double.compare(Double.parseDouble(av), Double.parseDouble(bv));
                } catch (Exception e) {
                  cmp = av.compareToIgnoreCase(bv);
                }
                return descending ? -cmp : cmp;
              });
          Map<Long, Workbook.Cell> sorted = new HashMap<>();
          for (int r = 0; r < order.size(); r++) {
            int old = order.get(r);
            for (int c = 0; c < sheet.lastCol(); c++) {
              Workbook.Cell cell = sheet.cell(old, c);
              if (cell != null) sorted.put(Workbook.Sheet.key(r, c), cell);
            }
          }
          sheet.cells.clear();
          sheet.cells.putAll(sorted);
        });
  }

  final class Grid extends View {
    private final Paint paint = new Paint(3);
    private final float cellW = dp(118), cellH = dp(41), headW = dp(48), headH = dp(33);
    private float scrollX, scrollY, downX, downY, lastX, lastY;
    private boolean dragged;

    Grid(Context context) {
      super(context);
      paint.setTypeface(Typeface.create("sans", Typeface.NORMAL));
    }

    void reset() {
      scrollX = scrollY = 0;
      invalidate();
    }

    void showCell() {
      scrollX = Math.max(0, selectedCol * cellW - cellW);
      scrollY = Math.max(0, selectedRow * cellH - cellH);
      invalidate();
    }

    @Override
    protected void onDraw(Canvas canvas) {
      Workbook.Sheet sheet = book.current();
      int width = getWidth(), height = getHeight();
      canvas.drawColor(Color.WHITE);
      paint.setStyle(Paint.Style.FILL);
      paint.setColor(0xFFF0F4F6);
      canvas.drawRect(0, 0, width, headH, paint);
      canvas.drawRect(0, 0, headW, height, paint);
      int firstCol = (int) (scrollX / cellW), firstRow = (int) (scrollY / cellH);
      int endCol = Math.min(sheet.cols, firstCol + (int) (width / cellW) + 2),
          endRow = Math.min(sheet.rows, firstRow + (int) (height / cellH) + 2);
      paint.setTextSize(dp(12));
      for (int col = firstCol; col < endCol; col++) {
        float left = headW + col * cellW - scrollX;
        paint.setColor(0xFF49616F);
        canvas.drawText(Workbook.columnName(col), left + dp(8), dp(22), paint);
      }
      for (int row = firstRow; row < endRow; row++) {
        float top = headH + row * cellH - scrollY;
        paint.setColor(0xFF49616F);
        canvas.drawText(Integer.toString(row + 1), dp(8), top + dp(25), paint);
        for (int col = firstCol; col < endCol; col++) {
          float left = headW + col * cellW - scrollX;
          Workbook.Cell cell = sheet.cell(row, col);
          paint.setColor(cell == null ? Color.WHITE : cell.fill);
          canvas.drawRect(left, top, left + cellW, top + cellH, paint);
          if (row == selectedRow && col == selectedCol) {
            paint.setColor(0xFFD3EDF7);
            paint.setStyle(Paint.Style.STROKE);
            paint.setStrokeWidth(dp(2));
            canvas.drawRect(left + 1, top + 1, left + cellW - 1, top + cellH - 1, paint);
            paint.setStyle(Paint.Style.FILL);
          }
          paint.setColor(0xFFD9E1E5);
          paint.setStrokeWidth(1);
          canvas.drawLine(left, top + cellH, left + cellW, top + cellH, paint);
          canvas.drawLine(left + cellW, top, left + cellW, top + cellH, paint);
          String value = sheet.display(row, col);
          paint.setColor(value.startsWith("#") ? 0xFFBD3131 : 0xFF1B303B);
          paint.setTypeface(
              cell != null && cell.bold
                  ? Typeface.create("sans", Typeface.BOLD)
                  : Typeface.create("sans", Typeface.NORMAL));
          canvas.save();
          canvas.clipRect(left + dp(5), top, left + cellW - dp(5), top + cellH);
          canvas.drawText(value, left + dp(8), top + dp(26), paint);
          canvas.restore();
        }
      }
    }

    @Override
    public boolean onTouchEvent(MotionEvent event) {
      switch (event.getActionMasked()) {
        case MotionEvent.ACTION_DOWN:
          downX = lastX = event.getX();
          downY = lastY = event.getY();
          dragged = false;
          return true;
        case MotionEvent.ACTION_MOVE:
          float dx = event.getX() - lastX, dy = event.getY() - lastY;
          if (Math.abs(event.getX() - downX) > dp(5) || Math.abs(event.getY() - downY) > dp(5))
            dragged = true;
          if (dragged) {
            scrollX =
                Math.max(
                    0, Math.min(book.current().cols * cellW - getWidth() + headW, scrollX - dx));
            scrollY =
                Math.max(
                    0, Math.min(book.current().rows * cellH - getHeight() + headH, scrollY - dy));
            invalidate();
          }
          lastX = event.getX();
          lastY = event.getY();
          return true;
        case MotionEvent.ACTION_UP:
          if (!dragged && event.getX() > headW && event.getY() > headH) {
            selectedCol =
                Math.min(book.current().cols - 1, (int) ((event.getX() - headW + scrollX) / cellW));
            selectedRow =
                Math.min(book.current().rows - 1, (int) ((event.getY() - headH + scrollY) / cellH));
            refresh();
            editCell();
          }
          return true;
        default:
          return true;
      }
    }
  }
}
