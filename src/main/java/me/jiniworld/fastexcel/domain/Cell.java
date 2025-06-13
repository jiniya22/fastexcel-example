package me.jiniworld.fastexcel.domain;

public record Cell(int colSize, int rowSize, CellType cellType, String value, CustomStyle style) {

  public Cell {
    if (colSize < 1 || rowSize < 1) {
      throw new IllegalArgumentException("길이는 최소 1 이상의 값이여야 합니다.");
    }
  }

  public static Cell of(int colSize, int rowSize, Object value) {
    return new Cell(colSize, rowSize, getCellType(value), convertString(value), null);
  }

  public static Cell of(int colSize, int rowSize, Object value, CustomStyle style) {
    return new Cell(colSize, rowSize, getCellType(value), convertString(value), style);
  }

  private static CellType getCellType(Object value) {
    return value instanceof Number ? CellType.NUMERIC : CellType.STRING;
  }

  private static String convertString(Object value) {
    return value == null ? null : value.toString();
  }

  public enum CellType {
    NUMERIC,
    STRING
  }
}
