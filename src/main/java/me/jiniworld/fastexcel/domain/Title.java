package me.jiniworld.fastexcel.domain;

public record Title(
    String title1,
    String title2,
    String title3,
    String title4,
    String title5,
    String title6,
    String title7,
    String title8,
    String title9,
    String title10,
    String title11,
    String title12,
    String title13,
    String title14,
    String title15,
    String title16) {
  public static Title sample() {
    return new Title(
        "t1", "t2", "t3", "t4", "t5", "t6", "t7", "t8", "t9", "t10", "t11", "t12", "t13", "t14",
        "t15", "t16");
  }
}
