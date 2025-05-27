package me.jiniworld.fastexcel.domain;

public record Content(
    String data1,
    String data2,
    String data3,
    String data4,
    String data5,
    String data6,
    String data7,
    String data8,
    String data9,
    String data10,
    String data11,
    String data12,
    String data13,
    String data14) {
  public static Content sample() {
    return new Content(
        "d1",
        "d2",
        "d3",
        "d4",
        "d5",
        "d6",
        "d7",
        "d8",
        "d9",
        "https://blog.jiniworld.me/",
        "d11",
        "d12",
        "d13",
        "d14");
  }
}
