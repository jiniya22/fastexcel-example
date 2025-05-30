package me.jiniworld.fastexcel.common;

public class ExcelCreateException extends RuntimeException {

  public ExcelCreateException() {
    super("엑셀 생성 중 오류 발생");
  }
}
