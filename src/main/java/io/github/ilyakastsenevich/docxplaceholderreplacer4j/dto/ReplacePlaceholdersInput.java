package io.github.ilyakastsenevich.docxplaceholderreplacer4j.dto;

import lombok.*;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Getter
@Setter
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class ReplacePlaceholdersInput {

  @Builder.Default
  private Map<String, List<TextValue>> placeholderToValueMap = new HashMap<>(5);

  private byte[] docxDocument;

  @Getter
  @Setter
  @Builder
  @AllArgsConstructor
  public static class TextValue {
    private String text;
    private TextValueFormat format;
  }

  @Getter
  @Setter
  @Builder
  @AllArgsConstructor
  public static class TextValueFormat {
    private Boolean bold;
    private String colorHex;
    private String fontFamily;
    private Integer fontSize;
    private String underlinePattern;
    private String underlineColorHex;
  }

  public void add(String placeholder, TextValue textValue) {
    add(placeholder, List.of(textValue));
  }

  public void add(String placeholder, List<TextValue> textValues) {
    placeholderToValueMap.put(placeholder, textValues);
  }

  public void add(String placeholder, String value) {
    add(placeholder, new TextValue(value, null));
  }

  public void clear() {
    placeholderToValueMap.clear();
    docxDocument = null;
  }
}