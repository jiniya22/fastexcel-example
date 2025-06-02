package me.jiniworld.fastexcel.domain;

import lombok.Builder;
import org.dhatim.fastexcel.Color;

@Builder
public record CustomStyle(
    String horizontalAlignment,
    String fillColor,
    String fontColor,
    Integer fontSize,
    boolean isBold,
    boolean isLink) {
  public static final String ALIGNMENT_CENTER = "center";

  public static CustomStyle titleStyle(CustomStyle customStyle) {
    final CustomStyle basicTitleStyle =
        CustomStyle.builder()
            .horizontalAlignment(ALIGNMENT_CENTER)
            .fillColor("002060")
            .fontSize(9)
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
    return CustomStyle.builder().horizontalAlignment(ALIGNMENT_CENTER).build();
  }

  /** 가운데 정렬 && 글씨체 두껍게 */
  public static CustomStyle style2() {
    return CustomStyle.builder().horizontalAlignment(ALIGNMENT_CENTER).isBold(true).build();
  }

  /** 가운데 정렬 && font색 red && 글씨체 두껍게 */
  public static CustomStyle style3() {
    return CustomStyle.builder()
        .horizontalAlignment(ALIGNMENT_CENTER)
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
}
