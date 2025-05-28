package me.jiniworld.fastexcel;

import lombok.RequiredArgsConstructor;
import me.jiniworld.fastexcel.service.CreateExcelService;
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
}
