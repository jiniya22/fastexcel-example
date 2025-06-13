package me.jiniworld.fastexcel.controller;

import lombok.RequiredArgsConstructor;
import me.jiniworld.fastexcel.service.CreateExcelService;
import org.springframework.http.HttpStatus;
import org.springframework.web.bind.annotation.*;

@RequiredArgsConstructor
@RequestMapping("/excel")
@RestController
public class ExcelController {

  private final CreateExcelService createExcelService;

  @ResponseStatus(HttpStatus.CREATED)
  @PostMapping("/type1")
  public void createExcel1() {
    createExcelService.createExcel1();
  }

  @ResponseStatus(HttpStatus.CREATED)
  @PostMapping("/type2")
  public void createExcel2() {
    createExcelService.createExcel2();
  }

  @ResponseStatus(HttpStatus.CREATED)
  @PostMapping("/type4")
  public void createExcel4() {
    createExcelService.createExcel4();
  }
}
