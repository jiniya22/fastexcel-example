package me.jiniworld.fastexcel.domain;

import lombok.Builder;
import org.dhatim.fastexcel.Color;

@Builder
public record CustomStyle(
    String horizontalAlignment,
    String fillColor,
    String fontColor,
    Integer fontSize,
    boolean isWrapText,
    boolean isBold,
    boolean isLink) {
  public static final String DEFAULT_ALIGNMENT = "center";
  public static final int DEFAULT_FONT_SIZE = 10;

  public static CustomStyle titleStyle(CustomStyle customStyle) {
    final CustomStyle basicTitleStyle =
        CustomStyle.builder()
            .horizontalAlignment(DEFAULT_ALIGNMENT)
            .fillColor("002060")
            .fontSize(DEFAULT_FONT_SIZE)
            .fontColor(Color.WHITE)
            .isBold(true)
            .build();
    if (customStyle == null) {
      return basicTitleStyle;
    }
    return CustomStyle.builder()
        .horizontalAlignment(
            customStyle.horizontalAlignment() == null
                ? basicTitleStyle.horizontalAlignment
                : customStyle.horizontalAlignment)
        .fillColor(
            customStyle.fillColor == null ? basicTitleStyle.fillColor : customStyle.fillColor)
        .fontColor(
            customStyle.fontColor == null ? basicTitleStyle.fontColor : customStyle.fontColor)
        .fontSize(customStyle.fontSize == null ? basicTitleStyle.fontSize : customStyle.fontSize)
        .isBold(customStyle.isBold())
        .isLink(customStyle.isLink())
        .build();
  }

  /** 가운데 정렬 */
  public static CustomStyle style1() {
    return CustomStyle.builder().horizontalAlignment(DEFAULT_ALIGNMENT).build();
  }

  /** 가운데 정렬 && 글씨체 두껍게 */
  public static CustomStyle style2() {
    return CustomStyle.builder().horizontalAlignment(DEFAULT_ALIGNMENT).isBold(true).build();
  }

  /** 가운데 정렬 && font색 red && 글씨체 두껍게 */
  public static CustomStyle style3() {
    return CustomStyle.builder()
        .horizontalAlignment(DEFAULT_ALIGNMENT)
        .fontColor(Color.RED)
        .isBold(true)
        .build();
  }

  /** 링크 타입 */
  public static CustomStyle style4() {
    return CustomStyle.builder().isLink(true).build();
  }

  /** font색 red */
  public static CustomStyle style5() {
    return CustomStyle.builder().fontColor(Color.RED).build();
  }

  /** 폰트 지정 && 글씨체 두껍게 */
  public static CustomStyle style6(int fontSize) {
    return CustomStyle.builder().isBold(true).fontSize(fontSize).build();
  }

  /** 연한회색 배경 && 글씨체 두껍게 */
  public static CustomStyle style7() {
    return CustomStyle.builder().fillColor("f2f2f2").isBold(true).fontColor(Color.BLACK).build();
  }

  /** font색 red && 글씨체 두껍게 */
  public static CustomStyle style8() {
    return CustomStyle.builder().fontColor(Color.RED).isBold(true).build();
  }

  /** 자동 줄바꿈 */
  public static CustomStyle wrapText() {
    return CustomStyle.builder().isWrapText(true).build();
  }
}
