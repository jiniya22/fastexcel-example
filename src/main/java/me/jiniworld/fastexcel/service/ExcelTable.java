package me.jiniworld.fastexcel.service;

import me.jiniworld.fastexcel.domain.Cell;
import me.jiniworld.fastexcel.domain.CustomStyle;
import org.dhatim.fastexcel.*;

public class ExcelTable {
  private static final int COLUMN_WIDTH = 12;

  private final Worksheet ws;
  private final int totalColSize;
  private final boolean autoLineWrap; //  자동 row 바꿈 (지정한 col 수를 채웠을 경우 다음 row로 자동 이동)
  private int top;
  private int left;
  private int bottom;
  private int right;
  private Cell tempCell; // autoLine() 에서 활용하기 위해 임시 저장하는 값

  public ExcelTable(Worksheet ws, int totalColSize, boolean autoLineWrap) {
    this.ws = ws;
    this.totalColSize = totalColSize;
    this.autoLineWrap = autoLineWrap;

    for (int i = 0; i < totalColSize; i++) {
      ws.width(i, COLUMN_WIDTH);
    }
  }

  public void applyTh(Cell cell) {
    commonStyleSetter(
            Cell.of(
                cell.colSize(), cell.rowSize(), cell.value(), CustomStyle.titleStyle(cell.style())))
        .set();
  }

  public void applyTd(Cell cell) {
    commonStyleSetter(cell).set();
  }

  private void applyStyle(Cell cell, StyleSetter style) {
    if (cell.style() == null) {
      return;
    }
    var customStyle = cell.style();
    if (customStyle.horizontalAlignment() != null) {
      style.horizontalAlignment(customStyle.horizontalAlignment());
    }
    if (customStyle.fillColor() != null) {
      style.fillColor(customStyle.fillColor());
    }
    if (customStyle.fontColor() != null) {
      style.fontColor(customStyle.fontColor());
    }
    if (customStyle.fontSize() != null) {
      style.fontSize(customStyle.fontSize());
    }
    if (customStyle.isBold()) {
      style.bold();
    }
    if (customStyle.isWrapText()) {
      style.wrapText(true);
    }
    if (customStyle.isLink()) {
      ws.hyperlink(top, left, new HyperLink(cell.value(), cell.value()));
      style.fontColor("0000FF").underlined();
    }
  }

  private StyleSetter commonStyleSetter(Cell cell) {
    tempCell = cell;
    bottom = top + cell.rowSize() - 1;
    right = left + cell.colSize() - 1;
    var result =
        ws.range(top, left, bottom, right)
            .style()
            .fontSize(CustomStyle.DEFAULT_FONT_SIZE)
            .borderColor("d4d4d4")
            .borderStyle(BorderStyle.THIN)
            .horizontalAlignment("left")
            .verticalAlignment(CustomStyle.DEFAULT_ALIGNMENT)
            .merge();
    if (Cell.CellType.NUMERIC == cell.cellType()) {
      ws.value(top, left, Long.parseLong(cell.value()));
    } else {
      ws.value(top, left, cell.value());
    }

    left += cell.colSize();
    if (this.autoLineWrap && right >= totalColSize - 1) {
      lineWrap();
    }
    applyStyle(cell, result);
    return result;
  }

  void lineWrap() {
    top += this.tempCell.rowSize();
    left = 0;
    bottom = top;
    right = 0;
  }
}
