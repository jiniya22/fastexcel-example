package me.jiniworld.fastexcel.service;

import lombok.RequiredArgsConstructor;
import me.jiniworld.fastexcel.domain.Cell;
import me.jiniworld.fastexcel.domain.CustomStyle;
import org.dhatim.fastexcel.*;

@RequiredArgsConstructor
public class ExcelTableService {
  private final Worksheet ws;
  private final int totalColSize;
  private int top;
  private int left;
  private int bottom;
  private int right;

  public void applyTh(Cell cell) {
    commonStyleSetter(
            new Cell(
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
    if (customStyle.isLink()) {
      ws.hyperlink(top, left, new HyperLink(cell.value(), cell.value()));
      style.fontColor("0000FF").underlined();
    }
  }

  private StyleSetter commonStyleSetter(Cell cell) {
    bottom = top + cell.rowSize() - 1;
    right = left + cell.colSize() - 1;
    var result =
        ws.range(top, left, bottom, right)
            .style()
            .borderColor("d4d4d4")
            .borderStyle(BorderStyle.THIN)
            .verticalAlignment(CustomStyle.ALIGNMENT_CENTER)
            .merge();
    ws.value(top, left, cell.value());
    left += cell.colSize();
    if (right >= totalColSize - 1) {
      top += cell.rowSize();
      left = 0;
      bottom = top;
      right = 0;
    }
    applyStyle(cell, result);
    return result;
  }
}
