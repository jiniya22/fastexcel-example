package me.jiniworld.fastexcel.service;

import java.io.OutputStream;
import lombok.experimental.UtilityClass;
import me.jiniworld.fastexcel.domain.CustomStyle;
import org.dhatim.fastexcel.*;

@UtilityClass
public class ExcelTableUtils {

  private final String TITLE_COLOR_BACKGROUND = "002060";

  public void applyTh(Worksheet ws, int row, int col, int rowspan, int colspan, String value) {
    titleStyleSetter(ws, row, col, rowspan, colspan).set();
    ws.value(row, col, value);
  }

  public void applyTd(Worksheet ws, int row, int col, int rowspan, int colspan, String value) {
    commonStyleSetter(ws, row, col, rowspan, colspan).set();
    ws.value(row, col, value);
  }

  public void applyTd(
      Worksheet ws,
      int row,
      int col,
      int rowspan,
      int colspan,
      String value,
      CustomStyle customStyle) {
    var style = commonStyleSetter(ws, row, col, rowspan, colspan);

    if (customStyle.fontColor() != null) {
      style.fontColor(customStyle.fontColor());
    }
    if (customStyle.horizontalAlignment() != null) {
      style.horizontalAlignment(customStyle.horizontalAlignment());
    }
    if (customStyle.isBold()) {
      style.bold();
    }
    if (customStyle.isLink()) {
      style.fontColor("0000FF").underlined();
      ws.hyperlink(row, col, new HyperLink(value, value));
    }
    style.set();
    ws.value(row, col, value);
  }

  public Workbook createWorkbook(OutputStream os) {
    Workbook workbook = new Workbook(os, "fastexcel-example", "1.0");
    workbook.setGlobalDefaultFont("Malgun Gothic", 10);
    return workbook;
  }

  public void initLayout(Worksheet ws, int totalColSize) {
    int columnWidth = 12;

    for (int i = 0; i < totalColSize; i++) {
      ws.width(i, columnWidth);
    }
  }

  private StyleSetter commonStyleSetter(Worksheet ws, int row, int col, int rowspan, int colspan) {
    System.out.printf(
        "%d %d %d %d%n",
        row, col, row + (rowspan > 0 ? rowspan - 1 : 0), col + (colspan > 0 ? colspan - 1 : 0));
    return ws.range(
            row, col, row + (rowspan > 0 ? rowspan - 1 : 0), col + (colspan > 0 ? colspan - 1 : 0))
        .style()
        .borderColor("d4d4d4")
        .borderStyle(BorderStyle.THIN)
        .verticalAlignment(CustomStyle.ALIGNMENT_CENTER)
        .merge();
  }

  public StyleSetter titleStyleSetter(Worksheet ws, int row, int col, int rowspan, int colspan) {
    return commonStyleSetter(ws, row, col, rowspan, colspan)
        .fillColor(TITLE_COLOR_BACKGROUND)
        .fontColor(Color.WHITE)
        .bold()
        .horizontalAlignment(CustomStyle.ALIGNMENT_CENTER);
  }
}
