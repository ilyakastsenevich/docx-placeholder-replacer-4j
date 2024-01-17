package io.github.ilyakastsenevich.docxplaceholderreplacer4j.replacer;


import io.github.ilyakastsenevich.docxplaceholderreplacer4j.dto.ReplacePlaceholdersInput;
import io.github.ilyakastsenevich.docxplaceholderreplacer4j.dto.ReplacePlaceholdersInput.TextValue;
import io.github.ilyakastsenevich.docxplaceholderreplacer4j.dto.ReplacePlaceholdersInput.TextValueFormat;
import io.github.ilyakastsenevich.docxplaceholderreplacer4j.exception.DocxPlaceholderReplacerException;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.*;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.*;
import java.util.stream.Collectors;

class DocxPlaceholderReplacerImpl implements DocxPlaceholderReplacer {

    @Override
    public byte[] replacePlaceholders(ReplacePlaceholdersInput input) {
        Map<String, List<TextValue>> placeholderToValueMap = input.getPlaceholderToValueMap();

        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
             ByteArrayInputStream templateInputStream = new ByteArrayInputStream(input.getDocxDocument());
             XWPFDocument xwpfDocument = new XWPFDocument(templateInputStream)) {

            replacePlaceholdersInHeaders(placeholderToValueMap, xwpfDocument);
            replacePlaceholdersInParagraphs(placeholderToValueMap, xwpfDocument);
            replacePlaceholderInTables(placeholderToValueMap, xwpfDocument);
            replacePlaceholdersInFooters(placeholderToValueMap, xwpfDocument);

            xwpfDocument.write(outputStream);

            return outputStream.toByteArray();
        } catch (Exception e) {
            throw new DocxPlaceholderReplacerException(e);
        }
    }

    private void replacePlaceholdersInHeaders(Map<String, List<TextValue>> dataParams, XWPFDocument xwpfDocument) {
        for (XWPFHeader header : xwpfDocument.getHeaderList()) {
            for (XWPFParagraph paragraph : header.getParagraphs()) {
                replacePlaceholdersInParagraph(dataParams, paragraph);
            }
        }
    }

    private void replacePlaceholdersInParagraphs(Map<String, List<TextValue>> dataParams, XWPFDocument xwpfDocument) {
        for (XWPFParagraph paragraph : xwpfDocument.getParagraphs()) {
            replacePlaceholdersInParagraph(dataParams, paragraph);
        }
    }

    private void replacePlaceholderInTables(Map<String, List<TextValue>> dataParams, XWPFDocument xwpfDocument) {
        for (XWPFTable table : xwpfDocument.getTables()) {
            for (XWPFTableRow tableRow : table.getRows()) {
                for (XWPFTableCell tableCell : tableRow.getTableCells()) {
                    for (XWPFParagraph paragraph : tableCell.getParagraphs()) {
                        replacePlaceholdersInParagraph(dataParams, paragraph);
                    }
                }
            }
        }
    }

    private void replacePlaceholdersInFooters(Map<String, List<TextValue>> dataParams, XWPFDocument xwpfDocument) {
        for (XWPFFooter footer : xwpfDocument.getFooterList()) {
            for (XWPFParagraph paragraph : footer.getParagraphs()) {
                replacePlaceholdersInParagraph(dataParams, paragraph);
            }
        }
    }

    private void replacePlaceholdersInParagraph(Map<String, List<TextValue>> dataParams, XWPFParagraph paragraph) {
        for (Map.Entry<String, List<TextValue>> entry : dataParams.entrySet()) {
            String originalText = paragraph.getText();

            if (StringUtils.containsIgnoreCase(originalText, entry.getKey()) && entry.getValue() != null) {

                TextValueFormat formatOfFirstRun = getFormatOfFirstRun(paragraph);

                // remove all text from current paragraph
                int runsMaxIndex = paragraph.getRuns().size() - 1;
                for (int i = runsMaxIndex; i >= 0; i--) {
                    paragraph.removeRun(i);
                }

                String textBeforePlaceholder = StringUtils.substringBefore(originalText, entry.getKey());
                String textAfterPlaceholder = StringUtils.substringAfter(originalText, entry.getKey());

                XWPFRun beforeRun = paragraph.createRun();
                beforeRun.setText(textBeforePlaceholder);
                applyFormat(beforeRun, formatOfFirstRun);

                List<TextValue> textValues = entry.getValue().stream().filter(Objects::nonNull).collect(Collectors.toList());

                for (TextValue dataInput : textValues) {
                    String inputText = dataInput.getText();

                    if (inputText == null) {
                        continue;
                    }

                    addLeadingLineBreaks(paragraph, inputText);

                    String inputTextWithoutLeadingAndTrailingLineBreaks = inputText.replaceAll("(^[\\r\\n]+|[\\r\\n]+$)", "");
                    Iterator<String> lines = inputTextWithoutLeadingAndTrailingLineBreaks.lines().iterator();

                    while (lines.hasNext()) {
                        String line = lines.next();

                        XWPFRun r = paragraph.createRun();
                        r.setText(line);
                        TextValueFormat inputFormat = dataInput.getFormat();
                        applyFormat(r, Optional.ofNullable(inputFormat).orElse(formatOfFirstRun));

                        if (lines.hasNext()) {
                            r.addCarriageReturn();
                        }
                    }

                    addTrailingLineBreaks(paragraph, inputText);
                }

                XWPFRun afterRun = paragraph.createRun();
                afterRun.setText(textAfterPlaceholder);
                applyFormat(afterRun, formatOfFirstRun);
            }
        }
    }


    private void addLeadingLineBreaks(XWPFParagraph paragraph, String input) {
        String inputTextWithoutWhiteSpaces = input.replaceAll(" ", "");

        //do not add line breaks if string contains only \n because it is added by addTrailingLineBreaks method
        if (inputTextWithoutWhiteSpaces.startsWith("\n") && !inputTextWithoutWhiteSpaces.chars().allMatch(c -> c == '\n')) {
            XWPFRun r = paragraph.createRun();
            for (int i = 0; i < inputTextWithoutWhiteSpaces.length(); i++) {
                if (inputTextWithoutWhiteSpaces.charAt(i) == '\n') {
                    r.addCarriageReturn();
                } else {
                    break;
                }
            }

        }
    }

    private void addTrailingLineBreaks(XWPFParagraph paragraph, String input) {
        String inputTextWithoutWhiteSpaces = input.replaceAll(" ", "");

        if (inputTextWithoutWhiteSpaces.endsWith("\n")) {
            XWPFRun r = paragraph.createRun();
            for (int i = inputTextWithoutWhiteSpaces.length() - 1; i >= 0; i--) {
                if (inputTextWithoutWhiteSpaces.charAt(i) == '\n') {
                    r.addCarriageReturn();
                } else {
                    break;
                }
            }

        }
    }

    private void applyFormat(XWPFRun r, TextValueFormat format) {
        if (r != null && format != null) {
            Optional.ofNullable(format.getBold()).ifPresent(r::setBold);
            Optional.ofNullable(format.getColorHex()).ifPresent(r::setColor);
            Optional.ofNullable(format.getFontFamily()).ifPresent(r::setFontFamily);
            Optional.ofNullable(format.getFontSize()).ifPresent(r::setFontSize);
            Optional.ofNullable(format.getUnderlinePattern()).ifPresent(pattern -> r.setUnderline(UnderlinePatterns.valueOf(pattern)));
            Optional.ofNullable(format.getUnderlineColorHex()).ifPresent(r::setUnderlineColor);
        }
    }

    private TextValueFormat getFormatOfFirstRun(XWPFParagraph paragraph) {
        return paragraph.getRuns()
                .stream()
                .findFirst().map(
                        run -> TextValueFormat.builder()
                                .bold(run.isBold())
                                .colorHex(run.getColor())
                                .fontFamily(run.getFontFamily())
                                .fontSize(Optional.ofNullable(run.getFontSizeAsDouble()).map(d -> d.intValue()).orElse(null))
                                .underlinePattern(run.getUnderline().name())
                                .underlineColorHex(run.getUnderlineColor())
                                .build())
                .orElse(null);
    }
}