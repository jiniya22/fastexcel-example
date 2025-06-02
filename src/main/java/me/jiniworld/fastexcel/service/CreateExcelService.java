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
        Workbook workbook = ExcelTableUtils.createWorkbook(os);
        Worksheet ws = workbook.newWorksheet("Sheet1")) {
      final int TOTAL_COL_SIZE = 10;
      ExcelTableUtils.initLayout(ws, TOTAL_COL_SIZE);

      ExcelTableUtils.titleStyleSetter(ws, 0, 0, 1, TOTAL_COL_SIZE).fontSize(18).set();
      ws.value(0, 0, "Title");

      ExcelTableUtils.applyTh(ws, 1, 0, 1, 1, title.title1());
      ExcelTableUtils.applyTd(ws, 1, 1, 1, 4, content.data1());
      ExcelTableUtils.applyTh(ws, 1, 5, 1, 1, title.title2());
      ExcelTableUtils.applyTd(ws, 1, 6, 1, 3, content.data2());
      ExcelTableUtils.applyTd(ws, 1, 9, 1, 1, content.data2());

      ExcelTableUtils.applyTh(ws, 2, 0, 1, 1, title.title3());
      ExcelTableUtils.applyTd(ws, 2, 1, 1, 9, content.data3());

      ExcelTableUtils.applyTh(ws, 3, 0, 1, 1, title.title4());
      ExcelTableUtils.applyTd(ws, 3, 1, 1, 9, content.data4());

      ExcelTableUtils.applyTh(ws, 4, 0, 1, 1, title.title5());
      ExcelTableUtils.applyTd(ws, 4, 1, 1, 4, content.data5());
      ExcelTableUtils.applyTh(ws, 4, 5, 1, 1, title.title6());
      ExcelTableUtils.applyTd(ws, 4, 6, 1, 4, content.data6());

      ExcelTableUtils.applyTh(ws, 5, 0, 1, 1, title.title7());
      ExcelTableUtils.applyTd(ws, 5, 1, 1, 4, content.data7());
      ExcelTableUtils.applyTh(ws, 5, 5, 1, 1, title.title8());
      ExcelTableUtils.applyTd(ws, 5, 6, 1, 3, content.data8());
      ExcelTableUtils.applyTd(ws, 5, 9, 1, 1, content.data8());

      ExcelTableUtils.applyTh(ws, 6, 0, 1, 1, title.title9());
      ExcelTableUtils.applyTd(ws, 6, 1, 1, 4, content.data9());
      ExcelTableUtils.applyTh(ws, 6, 5, 1, 1, title.title10());
      ExcelTableUtils.applyTd(ws, 6, 6, 1, 4, content.data10());

      ExcelTableUtils.applyTh(ws, 7, 0, 1, 1, title.title11());
      ExcelTableUtils.applyTd(ws, 7, 1, 1, 4, content.data11());
      ExcelTableUtils.applyTh(ws, 7, 5, 1, 1, title.title12());
      ExcelTableUtils.applyTd(ws, 7, 6, 1, 2, content.data12());
      ExcelTableUtils.applyTh(ws, 7, 8, 1, 1, title.title13());
      ExcelTableUtils.applyTd(ws, 7, 9, 1, 1, content.data13());

      ExcelTableUtils.applyTh(ws, 8, 0, 2, 1, title.title14());
      ExcelTableUtils.applyTd(ws, 8, 1, 2, 9, content.data14());

    } catch (IOException e) {
      throw new ExcelCreateException();
    }
  }

  public void createExcel3() {
    Title title = Title.sample();
    Content content = Content.sample();
    try (OutputStream os = new FileOutputStream("output3.xlsx");
        Workbook workbook = ExcelTableUtils.createWorkbook(os);
        Worksheet ws = workbook.newWorksheet("Sheet1")) {
      final int TOTAL_COL_SIZE = 8;
      ExcelTableUtils.initLayout(ws, TOTAL_COL_SIZE);

      ExcelTableService e = new ExcelTableService(ws, TOTAL_COL_SIZE);
      e.applyTh(Cell.of(TOTAL_COL_SIZE, 1, "Title", CustomStyle.style6(18)));

      e.applyTh(Cell.of(1, 1, title.title1()));
      e.applyTd(Cell.of(3, 1, content.data1()));
      e.applyTh(Cell.of(1, 1, title.title2()));
      e.applyTd(Cell.of(2, 1, content.data2()));
      e.applyTd(Cell.of(1, 1, content.data2()));

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

  public void createExcel4() {
    Title title = Title.sample();
    Content content = Content.sample();
    try (OutputStream os = new FileOutputStream("output4.xlsx");
        Workbook workbook = ExcelTableUtils.createWorkbook(os);
        Worksheet ws = workbook.newWorksheet("Sheet1")) {
      final int TOTAL_COL_SIZE = 8;
      ExcelTableUtils.initLayout(ws, TOTAL_COL_SIZE);

      ExcelTableUtils.titleStyleSetter(ws, 0, 0, 1, TOTAL_COL_SIZE).fontSize(18).set();
      ws.value(0, 0, "Title");

      ExcelTableUtils.applyTh(ws, 1, 0, 1, 1, title.title1());
      ExcelTableUtils.applyTd(ws, 1, 1, 1, 3, content.data1());
      ExcelTableUtils.applyTh(ws, 1, 4, 1, 1, title.title2());
      ExcelTableUtils.applyTd(ws, 1, 5, 1, 2, content.data2());
      ExcelTableUtils.applyTd(ws, 1, 7, 1, 1, content.data2());

      ExcelTableUtils.applyTh(ws, 2, 0, 1, 1, title.title3());
      ExcelTableUtils.applyTd(ws, 2, 1, 1, 3, content.data3());
      ExcelTableUtils.applyTh(ws, 2, 4, 1, 1, title.title4());
      ExcelTableUtils.applyTd(ws, 2, 5, 1, 3, content.data4());

      ExcelTableUtils.applyTh(ws, 3, 0, 1, 1, title.title5());
      ExcelTableUtils.applyTd(ws, 3, 1, 1, 3, content.data5());
      ExcelTableUtils.applyTh(ws, 3, 4, 1, 1, title.title6());
      ExcelTableUtils.applyTd(ws, 3, 5, 1, 3, content.data6());

      ExcelTableUtils.applyTh(ws, 4, 0, 1, 1, title.title7());
      ExcelTableUtils.applyTd(ws, 4, 1, 1, 7, content.data7());

      ExcelTableUtils.applyTh(ws, 5, 0, 1, 1, title.title8());
      ExcelTableUtils.applyTd(ws, 5, 1, 1, 3, content.data8());
      ExcelTableUtils.applyTh(ws, 5, 4, 1, 1, title.title9());
      ExcelTableUtils.applyTd(ws, 5, 5, 1, 3, content.data9());

      ExcelTableUtils.applyTh(ws, 6, 0, 1, 1, title.title10());
      ExcelTableUtils.applyTd(ws, 6, 1, 1, 3, content.data10());
      ExcelTableUtils.applyTh(ws, 6, 4, 1, 1, title.title11());
      ExcelTableUtils.applyTd(ws, 6, 5, 1, 3, content.data11());

      ExcelTableUtils.applyTh(ws, 7, 0, 1, 1, title.title12());
      ExcelTableUtils.applyTd(ws, 7, 1, 1, 3, content.data12());
      ExcelTableUtils.applyTh(ws, 7, 4, 1, 1, title.title13());
      ExcelTableUtils.applyTd(ws, 7, 5, 1, 3, content.data13());

      ExcelTableUtils.applyTh(ws, 8, 0, 1, 1, title.title14());
      ExcelTableUtils.applyTd(ws, 8, 1, 1, 3, content.data14(), CustomStyle.style4());
      ExcelTableUtils.applyTh(ws, 8, 4, 1, 1, title.title15());
      ExcelTableUtils.applyTd(ws, 8, 5, 1, 1, content.data15());
      ExcelTableUtils.applyTh(ws, 8, 6, 1, 1, title.title16());
      ExcelTableUtils.applyTd(ws, 8, 7, 1, 1, content.data16());

      ExcelTableUtils.applyTh(ws, 9, 0, 1, TOTAL_COL_SIZE, title.title17());
      ExcelTableUtils.applyTd(ws, 10, 0, 1, TOTAL_COL_SIZE, content.data17(), CustomStyle.style1());
      ExcelTableUtils.applyTd(ws, 11, 0, 1, TOTAL_COL_SIZE, content.data17(), CustomStyle.style1());
      ExcelTableUtils.applyTd(ws, 12, 0, 1, TOTAL_COL_SIZE, content.data17(), CustomStyle.style2());
      ExcelTableUtils.applyTd(ws, 13, 0, 1, TOTAL_COL_SIZE, content.data17(), CustomStyle.style3());

      ExcelTableUtils.applyTh(ws, 14, 0, 1, TOTAL_COL_SIZE, title.title18());

      ExcelTableUtils.applyTd(ws, 15, 0, 1, 1, title.title19());
      ExcelTableUtils.applyTd(ws, 15, 1, 1, 2, null);
      ExcelTableUtils.applyTd(ws, 15, 3, 1, 3, null);
      ExcelTableUtils.applyTd(ws, 15, 6, 1, 2, null);

      ExcelTableUtils.applyTd(ws, 16, 0, 1, 1, title.title20());
      ExcelTableUtils.applyTd(ws, 16, 1, 1, 2, null);
      ExcelTableUtils.applyTd(ws, 16, 3, 1, 3, null);
      ExcelTableUtils.applyTd(ws, 16, 6, 1, 2, null);
    } catch (IOException e) {
      throw new ExcelCreateException();
    }
  }
}
