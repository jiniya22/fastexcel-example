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
  @PostMapping("/type4")
  public void createExcel4() {
    createExcelService.createExcel4();
  }
}
