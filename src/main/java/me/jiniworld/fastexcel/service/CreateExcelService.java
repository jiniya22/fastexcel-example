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

  public void createExcel1() {
    Title title = Title.sample();
    Content content = Content.sample();
    try (OutputStream os = new FileOutputStream("output1.xlsx");
        Workbook workbook = ExcelTableUtils.createWorkbook(os);
        Worksheet ws = workbook.newWorksheet("Sheet1")) {
      int totalColSize = 10;
      ExcelTableUtils.initLayout(ws, totalColSize);

      ExcelTableUtils.titleStyleSetter(ws, 0, 0, totalColSize).fontSize(18).set();
      ws.value(0, 0, "Title");

      ExcelTableUtils.applyTh(ws, 1, 0, 1, title.title1());
      ExcelTableUtils.applyTd(ws, 1, 1, 4, content.data1());
      ExcelTableUtils.applyTh(ws, 1, 5, 1, title.title2());
      ExcelTableUtils.applyTd(ws, 1, 6, 3, content.data2());
      ExcelTableUtils.applyTd(ws, 1, 9, 1, content.data2());

      ExcelTableUtils.applyTh(ws, 2, 0, 1, title.title3());
      ExcelTableUtils.applyTd(ws, 2, 1, 9, content.data3());

      ExcelTableUtils.applyTh(ws, 3, 0, 1, title.title4());
      ExcelTableUtils.applyTd(ws, 3, 1, 9, content.data4());

      ExcelTableUtils.applyTh(ws, 4, 0, 1, title.title5());
      ExcelTableUtils.applyTd(ws, 4, 1, 4, content.data5());
      ExcelTableUtils.applyTh(ws, 4, 5, 1, title.title6());
      ExcelTableUtils.applyTd(ws, 4, 6, 4, content.data6());

      ExcelTableUtils.applyTh(ws, 5, 0, 1, title.title7());
      ExcelTableUtils.applyTd(ws, 5, 1, 4, content.data7());
      ExcelTableUtils.applyTh(ws, 5, 5, 1, title.title8());
      ExcelTableUtils.applyTd(ws, 5, 6, 3, content.data8());
      ExcelTableUtils.applyTd(ws, 5, 9, 1, content.data8());

      ExcelTableUtils.applyTh(ws, 6, 0, 1, title.title9());
      ExcelTableUtils.applyTd(ws, 6, 1, 4, content.data9());
      ExcelTableUtils.applyTh(ws, 6, 5, 1, title.title10());
      ExcelTableUtils.applyTd(ws, 6, 6, 4, content.data10());

      ExcelTableUtils.applyTh(ws, 7, 0, 1, title.title11());
      ExcelTableUtils.applyTd(ws, 7, 1, 4, content.data11());
      ExcelTableUtils.applyTh(ws, 7, 5, 1, title.title12());
      ExcelTableUtils.applyTd(ws, 7, 6, 2, content.data12());
      ExcelTableUtils.applyTh(ws, 7, 8, 1, title.title13());
      ExcelTableUtils.applyTd(ws, 7, 9, 1, content.data13());

    } catch (IOException e) {
      throw new RuntimeException(e);
    }
  }

  public void createExcel4() {
    Title title = Title.sample();
    Content content = Content.sample();
    try (OutputStream os = new FileOutputStream("output4.xlsx");
        Workbook workbook = ExcelTableUtils.createWorkbook(os);
        Worksheet ws = workbook.newWorksheet("Sheet1")) {
      int totalColSize = 8;
      ExcelTableUtils.initLayout(ws, totalColSize);

      ExcelTableUtils.titleStyleSetter(ws, 0, 0, totalColSize).fontSize(18).set();
      ws.value(0, 0, "Title");

      ExcelTableUtils.applyTh(ws, 1, 0, 1, title.title1());
      ExcelTableUtils.applyTd(ws, 1, 1, 3, content.data1());
      ExcelTableUtils.applyTh(ws, 1, 4, 1, title.title2());
      ExcelTableUtils.applyTd(ws, 1, 5, 2, content.data2());
      ExcelTableUtils.applyTd(ws, 1, 7, 1, content.data2());

      ExcelTableUtils.applyTh(ws, 2, 0, 1, title.title3());
      ExcelTableUtils.applyTd(ws, 2, 1, 7, content.data3());

      ExcelTableUtils.applyTh(ws, 3, 0, 1, title.title4());
      ExcelTableUtils.applyTd(ws, 3, 1, 3, content.data4());
      ExcelTableUtils.applyTh(ws, 3, 4, 1, title.title5());
      ExcelTableUtils.applyTd(ws, 3, 5, 3, content.data5());

      ExcelTableUtils.applyTh(ws, 4, 0, 1, title.title6());
      ExcelTableUtils.applyTd(ws, 4, 1, 3, content.data6());
      ExcelTableUtils.applyTh(ws, 4, 4, 1, title.title7());
      ExcelTableUtils.applyTd(ws, 4, 5, 3, content.data7());

      ExcelTableUtils.applyTh(ws, 5, 0, 1, title.title8());
      ExcelTableUtils.applyTd(ws, 5, 1, 3, content.data8());
      ExcelTableUtils.applyTh(ws, 5, 4, 1, title.title9());
      ExcelTableUtils.applyTd(ws, 5, 5, 3, content.data9());

      ExcelTableUtils.applyTh(ws, 6, 0, 1, title.title10());
      ExcelTableUtils.applyCustomTd(ws, 6, 1, 3, content.data10(), CustomStyle.style4());
      ExcelTableUtils.applyTh(ws, 6, 4, 1, title.title11());
      ExcelTableUtils.applyTd(ws, 6, 5, 1, content.data11());
      ExcelTableUtils.applyTh(ws, 6, 6, 1, title.title12());
      ExcelTableUtils.applyTd(ws, 6, 7, 1, content.data12());

      ExcelTableUtils.applyTh(ws, 7, 0, totalColSize, title.title13());
      ExcelTableUtils.applyCustomTd(ws, 8, 0, totalColSize, content.data13(), CustomStyle.style1());
      ExcelTableUtils.applyCustomTd(ws, 9, 0, totalColSize, content.data13(), CustomStyle.style1());
      ExcelTableUtils.applyCustomTd(
          ws, 10, 0, totalColSize, content.data13(), CustomStyle.style2());
      ExcelTableUtils.applyCustomTd(
          ws, 11, 0, totalColSize, content.data13(), CustomStyle.style3());

      ExcelTableUtils.applyTh(ws, 12, 0, 8, title.title14());
      ExcelTableUtils.applyTd(ws, 13, 0, 1, title.title15());
      ExcelTableUtils.applyTd(ws, 13, 1, 2, null);
      ExcelTableUtils.applyTd(ws, 13, 3, 3, null);
      ExcelTableUtils.applyTd(ws, 13, 6, 2, null);

      ExcelTableUtils.applyTd(ws, 14, 0, 1, title.title16());
      ExcelTableUtils.applyTd(ws, 14, 1, 2, null);
      ExcelTableUtils.applyTd(ws, 14, 3, 3, null);
      ExcelTableUtils.applyTd(ws, 14, 6, 2, null);

    } catch (IOException e) {
      throw new RuntimeException(e);
    }
  }
}
