package me.jiniworld.fastexcel.domain;

public record Cell(int colSize, int rowSize, String value, CustomStyle style) {

  public Cell {
    if (colSize < 1 || rowSize < 1) {
      throw new IllegalArgumentException("길이는 최소 1 이상의 값이여야 합니다.");
    }
  }

  public static Cell of(int colSize, int rowSize, String value) {
    return new Cell(colSize, rowSize, value, null);
  }

  public static Cell of(int colSize, int rowSize, String value, CustomStyle style) {
    return new Cell(colSize, rowSize, value, style);
  }
}
