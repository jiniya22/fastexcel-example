package me.jiniworld.fastexcel;

import lombok.RequiredArgsConstructor;
import me.jiniworld.fastexcel.service.CreateExcelService;
import org.springframework.http.HttpStatus;
import org.springframework.web.bind.annotation.*;

@RequiredArgsConstructor
@RequestMapping("")
@RestController
public class MainController {

  private final CreateExcelService createExcelService;

  @GetMapping("")
  public String heathCheck() {
    return "fastexcel-example is running";
  }

  @ResponseStatus(HttpStatus.CREATED)
  @PostMapping("/excel")
  public void createExcel() {
    createExcelService.createExcel();
  }
}
