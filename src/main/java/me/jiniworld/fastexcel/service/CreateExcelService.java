package me.jiniworld.fastexcel.service;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import me.jiniworld.fastexcel.common.ExcelCreateException;
import me.jiniworld.fastexcel.domain.Cell;
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
        Workbook workbook = createWorkbook(os);
        Worksheet ws = workbook.newWorksheet("Sheet1")) {
      final int TOTAL_COL_SIZE = 10;

      ExcelTable e = new ExcelTable(ws, TOTAL_COL_SIZE, true);
      e.applyTh(Cell.of(TOTAL_COL_SIZE, 1, "Title", CustomStyle.style6(18)));

      e.applyTh(Cell.of(1, 1, title.title1()));
      e.applyTd(Cell.of(4, 1, content.data1()));
      e.applyTh(Cell.of(1, 1, title.title2()));
      e.applyTd(Cell.of(3, 1, content.data2()));
      e.applyTd(Cell.of(1, 1, content.data2()));

      e.applyTh(Cell.of(1, 1, title.title3()));
      e.applyTd(Cell.of(9, 1, content.data3()));

      e.applyTh(Cell.of(1, 1, title.title4()));
      e.applyTd(Cell.of(9, 1, content.data4()));

      e.applyTh(Cell.of(1, 1, title.title5()));
      e.applyTd(Cell.of(4, 1, content.data5()));
      e.applyTh(Cell.of(1, 1, title.title6()));
      e.applyTd(Cell.of(4, 1, content.data6()));

      e.applyTh(Cell.of(1, 1, title.title7()));
      e.applyTd(Cell.of(4, 1, content.data7()));
      e.applyTh(Cell.of(1, 1, title.title8()));
      e.applyTd(Cell.of(3, 1, content.data8()));
      e.applyTd(Cell.of(1, 1, content.data8()));

      e.applyTh(Cell.of(1, 1, title.title9()));
      e.applyTd(Cell.of(4, 1, content.data9()));
      e.applyTh(Cell.of(1, 1, title.title10()));
      e.applyTd(Cell.of(4, 1, content.data10()));

      e.applyTh(Cell.of(1, 1, title.title11()));
      e.applyTd(Cell.of(4, 1, content.data11()));
      e.applyTh(Cell.of(1, 1, title.title12()));
      e.applyTd(Cell.of(2, 1, content.data12()));
      e.applyTh(Cell.of(1, 1, title.title13()));
      e.applyTd(Cell.of(1, 1, content.data13()));

      e.applyTh(Cell.of(1, 2, title.title14()));
      e.applyTd(Cell.of(9, 2, content.data14()));

      e.applyTh(Cell.of(TOTAL_COL_SIZE, 1, title.title15()));
      e.applyTh(Cell.of(1, 1, "Track", CustomStyle.style7()));
      e.applyTh(Cell.of(1, 1, "타이틀", CustomStyle.style7()));
      e.applyTh(Cell.of(3, 1, "곡명", CustomStyle.style7()));
      e.applyTh(Cell.of(2, 1, "곡 코드", CustomStyle.style7()));
      e.applyTh(Cell.of(3, 1, "아티스트명", CustomStyle.style7()));
      e.applyTd(Cell.of(1, 1, null));
      e.applyTd(Cell.of(1, 1, null));
      e.applyTd(Cell.of(3, 1, null));
      e.applyTd(Cell.of(2, 1, null));
      e.applyTd(Cell.of(3, 1, null));

      e.applyTh(Cell.of(TOTAL_COL_SIZE, 1, "info"));

      e.applyTh(Cell.of(1, 1, title.title1()));
      e.applyTd(Cell.of(TOTAL_COL_SIZE - 1, 1, content.data1()));

      e.applyTh(Cell.of(1, 1, title.title2()));
      e.applyTd(Cell.of(4, 1, content.data2()));
      e.applyTh(Cell.of(1, 1, title.title3()));
      e.applyTd(Cell.of(4, 1, content.data3()));

      e.applyTh(Cell.of(1, 1, title.title4()));
      e.applyTd(Cell.of(4, 1, content.data4()));
      e.applyTh(Cell.of(1, 1, title.title5()));
      e.applyTd(Cell.of(4, 1, content.data5()));

      e.applyTh(Cell.of(1, 1, title.title6()));
      e.applyTd(Cell.of(4, 1, content.data6()));
      e.applyTh(Cell.of(1, 1, title.title7()));
      e.applyTd(Cell.of(4, 1, content.data7()));

      e.applyTh(Cell.of(TOTAL_COL_SIZE, 1, title.title16()));
      for (int i = 0; i < 8; i++) {
        e.applyTd(Cell.of(5, 1, content.data16()));
        e.applyTd(Cell.of(5, 1, content.data16()));
      }

      e.applyTh(Cell.of(TOTAL_COL_SIZE, 1, title.title16()));

      e.applyTd(Cell.of(1, 1, title.title19(), CustomStyle.style1()));
      e.applyTd(Cell.of(2, 1, null));
      e.applyTd(Cell.of(4, 1, null));
      e.applyTd(Cell.of(3, 1, null));

      e.applyTd(Cell.of(1, 1, title.title20(), CustomStyle.style1()));
      e.applyTd(Cell.of(2, 1, null));
      e.applyTd(Cell.of(4, 1, null));
      e.applyTd(Cell.of(3, 1, null));
    } catch (IOException e) {
      throw new ExcelCreateException();
    }
  }

  public void createExcel4() {
    Title title = Title.sample();
    Content content = Content.sample();
    try (OutputStream os = new FileOutputStream("output4.xlsx");
        Workbook workbook = createWorkbook(os);
        Worksheet ws = workbook.newWorksheet("Sheet1")) {
      final int TOTAL_COL_SIZE = 8;
      ExcelTable e = new ExcelTable(ws, TOTAL_COL_SIZE, true);
      e.applyTh(Cell.of(TOTAL_COL_SIZE, 1, "Title", CustomStyle.style6(18)));

      e.applyTh(Cell.of(1, 1, title.title1()));
      e.applyTd(Cell.of(3, 1, content.data1()));
      e.applyTh(Cell.of(1, 1, title.title2()));
      e.applyTd(Cell.of(2, 1, content.data2()));
      e.applyTd(Cell.of(1, 1, content.data2()));

      e.applyTh(Cell.of(1, 1, title.title3()));
      e.applyTd(Cell.of(3, 1, content.data3()));
      e.applyTh(Cell.of(1, 1, title.title4()));
      e.applyTd(Cell.of(3, 1, content.data4()));

      e.applyTh(Cell.of(1, 1, title.title5()));
      e.applyTd(Cell.of(3, 1, content.data5()));
      e.applyTh(Cell.of(1, 1, title.title6()));
      e.applyTd(Cell.of(3, 1, content.data6()));

      e.applyTh(Cell.of(1, 1, title.title7()));
      e.applyTd(Cell.of(7, 1, content.data7()));

      e.applyTh(Cell.of(1, 1, title.title8()));
      e.applyTd(Cell.of(3, 1, content.data8()));
      e.applyTh(Cell.of(1, 1, title.title9()));
      e.applyTd(Cell.of(3, 1, content.data9()));

      e.applyTh(Cell.of(1, 1, title.title10()));
      e.applyTd(Cell.of(3, 1, content.data10()));
      e.applyTh(Cell.of(1, 1, title.title11()));
      e.applyTd(Cell.of(3, 1, content.data11()));

      e.applyTh(Cell.of(1, 1, title.title12()));
      e.applyTd(Cell.of(3, 1, content.data12()));
      e.applyTh(Cell.of(1, 1, title.title13()));
      e.applyTd(Cell.of(3, 1, content.data13()));

      e.applyTh(Cell.of(1, 1, title.title14()));
      e.applyTd(Cell.of(3, 1, content.data14(), CustomStyle.style4()));
      e.applyTh(Cell.of(1, 1, title.title15()));
      e.applyTd(Cell.of(1, 1, content.data15()));
      e.applyTh(Cell.of(1, 1, title.title16()));
      e.applyTd(Cell.of(1, 1, content.data16()));

      e.applyTh(Cell.of(TOTAL_COL_SIZE, 1, title.title17()));
      e.applyTd(Cell.of(TOTAL_COL_SIZE, 1, content.data17(), CustomStyle.style1()));
      e.applyTd(Cell.of(TOTAL_COL_SIZE, 1, content.data17(), CustomStyle.style1()));
      e.applyTd(Cell.of(TOTAL_COL_SIZE, 1, content.data17(), CustomStyle.style2()));
      e.applyTd(Cell.of(TOTAL_COL_SIZE, 1, content.data17(), CustomStyle.style3()));

      e.applyTh(Cell.of(TOTAL_COL_SIZE, 1, title.title18()));

      e.applyTd(Cell.of(1, 1, title.title19()));
      e.applyTd(Cell.of(2, 1, null));
      e.applyTd(Cell.of(3, 1, null));
      e.applyTd(Cell.of(2, 1, null));

      e.applyTd(Cell.of(1, 1, title.title20()));
      e.applyTd(Cell.of(2, 1, null));
      e.applyTd(Cell.of(3, 1, null));
      e.applyTd(Cell.of(2, 1, null));
    } catch (IOException e) {
      throw new ExcelCreateException();
    }
  }

  public Workbook createWorkbook(OutputStream os) {
    Workbook workbook = new Workbook(os, "fastexcel-example", "1.0");
    workbook.setGlobalDefaultFont("Malgun Gothic", 10);
    return workbook;
  }
}
