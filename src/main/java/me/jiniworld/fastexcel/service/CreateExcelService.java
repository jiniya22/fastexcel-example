package me.jiniworld.fastexcel.service;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import me.jiniworld.fastexcel.domain.Content;
import me.jiniworld.fastexcel.domain.CustomStyle;
import me.jiniworld.fastexcel.domain.Title;
import org.dhatim.fastexcel.*;
import org.springframework.stereotype.Service;

@Service
public class CreateExcelService {

  private final String TITLE_COLOR_BACKGROUND = "002060";
  private final String TITLE_COLOR_FONT = Color.WHITE;

  public void createExcel() {
    Title title = Title.sample();
    Content content = Content.sample();
    try (OutputStream os = new FileOutputStream("output.xlsx");
        Workbook workbook = createWorkbook(os);
        Worksheet ws = workbook.newWorksheet("Sheet1")) {
      initLayout(ws);

      int row = 0;
      int col = 0;

      applyTitleStyle(ws, row, col, 8).fontSize(18).set();
      ws.value(row, col, "Title");

      applyTh(ws, 1, 0, 1, title.title1());
      applyTd(ws, 1, 1, 3, content.data1());
      applyTh(ws, 1, 4, 1, title.title2());
      applyTd(ws, 1, 5, 2, content.data2());
      applyTd(ws, 1, 7, 1, content.data2());

      applyTh(ws, 2, 0, 1, title.title3());
      applyTd(ws, 2, 1, 7, content.data3());

      applyTh(ws, 3, 0, 1, title.title4());
      applyTd(ws, 3, 1, 3, content.data4());
      applyTh(ws, 3, 4, 1, title.title5());
      applyTd(ws, 3, 5, 3, content.data5());

      applyTh(ws, 4, 0, 1, title.title6());
      applyTd(ws, 4, 1, 3, content.data6());
      applyTh(ws, 4, 4, 1, title.title7());
      applyTd(ws, 4, 5, 3, content.data7());

      applyTh(ws, 5, 0, 1, title.title8());
      applyTd(ws, 5, 1, 3, content.data8());
      applyTh(ws, 5, 4, 1, title.title9());
      applyTd(ws, 5, 5, 3, content.data9());

      applyTh(ws, 6, 0, 1, title.title10());
      applyCustomTd(ws, 6, 1, 3, content.data10(), CustomStyle.style4());
      applyTh(ws, 6, 4, 1, title.title11());
      applyTd(ws, 6, 5, 1, content.data11());
      applyTh(ws, 6, 6, 1, title.title12());
      applyTd(ws, 6, 7, 1, content.data12());

      applyTh(ws, 7, 0, 8, title.title13());
      applyCustomTd(ws, 8, 0, 8, content.data13(), CustomStyle.style1());
      applyCustomTd(ws, 9, 0, 8, content.data13(), CustomStyle.style1());
      applyCustomTd(ws, 10, 0, 8, content.data13(), CustomStyle.style2());
      applyCustomTd(ws, 11, 0, 8, content.data13(), CustomStyle.style3());

      applyTh(ws, 12, 0, 8, title.title14());
      applyTd(ws, 13, 0, 1, title.title15());
      applyTd(ws, 13, 1, 2, null);
      applyTd(ws, 13, 3, 3, null);
      applyTd(ws, 13, 6, 2, null);

      applyTd(ws, 14, 0, 1, title.title16());
      applyTd(ws, 14, 1, 2, null);
      applyTd(ws, 14, 3, 3, null);
      applyTd(ws, 14, 6, 2, null);

    } catch (IOException e) {
      throw new RuntimeException(e);
    }
  }

  private StyleSetter applyCommonStyle(Worksheet ws, int row, int col, int colspan) {
    return ws.range(row, col, row, col + (colspan > 0 ? colspan - 1 : 0))
        .style()
        .borderColor("d4d4d4")
        .borderStyle(BorderStyle.THIN)
        .verticalAlignment(CustomStyle.ALIGNMENT_CENTER)
        .merge();
  }

  private StyleSetter applyTitleStyle(Worksheet ws, int row, int col, int colspan) {
    return applyCommonStyle(ws, row, col, colspan)
        .fillColor(TITLE_COLOR_BACKGROUND)
        .fontColor(Color.WHITE)
        .bold()
        .horizontalAlignment(CustomStyle.ALIGNMENT_CENTER);
  }

  private void applyTh(Worksheet ws, int row, int col, int colspan, String value) {
    applyTitleStyle(ws, row, col, colspan).set();
    ws.value(row, col, value);
  }

  private void applyTd(Worksheet ws, int row, int col, int colspan, String value) {
    applyCommonStyle(ws, row, col, colspan).set();
    ws.value(row, col, value);
  }

  private void applyCustomTd(
      Worksheet ws, int row, int col, int colspan, String value, CustomStyle customStyle) {
    var style = applyCommonStyle(ws, row, col, colspan);

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

  private Workbook createWorkbook(OutputStream os) {
    Workbook workbook = new Workbook(os, "fastexcel-example", "1.0");
    workbook.setGlobalDefaultFont("Malgun Gothic", 10);
    return workbook;
  }

  private void initLayout(Worksheet ws) {
    int columnWidth = 12;

    ws.width(0, columnWidth);
    ws.width(1, columnWidth);
    ws.width(2, columnWidth);
    ws.width(3, columnWidth);
    ws.width(4, columnWidth);
    ws.width(5, columnWidth);
    ws.width(6, columnWidth);
    ws.width(7, columnWidth);
  }
}
