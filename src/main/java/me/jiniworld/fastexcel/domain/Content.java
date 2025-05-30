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
    String data14,
    String data15,
    String data16,
    String data17,
    String data18) {
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
        "d10",
        "d11",
        "d12",
        "d13",
        "https://blog.jiniworld.me/",
        "d15",
        "d16",
        "d17",
        "d18");
  }
}
