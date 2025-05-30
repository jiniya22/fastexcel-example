package me.jiniworld.fastexcel.common;

import java.util.Map;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.RestControllerAdvice;
import org.springframework.web.servlet.mvc.method.annotation.ResponseEntityExceptionHandler;

@Slf4j
@RestControllerAdvice
@RequiredArgsConstructor
public class GlobalExceptionHandler extends ResponseEntityExceptionHandler {

  @ExceptionHandler({ExcelCreateException.class})
  protected ResponseEntity<Map<String, Object>> handleException(RuntimeException e) {
    return ResponseEntity.status(HttpStatus.BAD_REQUEST)
        .body(Map.of("code", HttpStatus.BAD_REQUEST.value(), "errorMessage", e.getMessage()));
  }
}
