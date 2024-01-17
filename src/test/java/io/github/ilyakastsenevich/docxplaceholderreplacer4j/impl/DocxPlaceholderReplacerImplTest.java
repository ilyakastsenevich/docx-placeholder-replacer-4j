package io.github.ilyakastsenevich.docxplaceholderreplacer4j.impl;

import io.github.ilyakastsenevich.docxplaceholderreplacer4j.dto.ReplacePlaceholdersInput;
import io.github.ilyakastsenevich.docxplaceholderreplacer4j.dto.ReplacePlaceholdersInput.TextValue;
import io.github.ilyakastsenevich.docxplaceholderreplacer4j.replacer.DocxPlaceholderReplacer;
import org.junit.jupiter.api.Test;

import java.io.IOException;

import static io.github.ilyakastsenevich.docxplaceholderreplacer4j.dto.ReplacePlaceholdersInput.TextValueFormat;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

class DocxPlaceholderReplacerImplTest {

    @Test
    void replacePlaceholders() throws IOException {
        //your docx template byte[]
        byte[] docxDocument = this.getClass().getClassLoader().getResourceAsStream("test_placeholders.docx").readAllBytes();

        //create new input object
        ReplacePlaceholdersInput input = new ReplacePlaceholdersInput();

        //add docx document
        input.setDocxDocument(docxDocument);

        //add simple text value replacement
        input.add("${header}", "replaced header");
        input.add("${footer}", "replaced footer");
        input.add("${simple1}", "replaced 1");
        input.add("${simple2}", "replaced 2");
        input.add("${simple3}", "replaced 3");
        input.add("${simple4}", "replaced 4");
        input.add("${keepFormat}", "replaced and original format preserved");

        //add text with formatting
        //create format object, every setting is optional
        TextValueFormat format = TextValueFormat.builder()
                .bold(true)
                .colorHex("008000") //green
                .fontFamily("Comic Sans MS")
                .fontSize(15)
                .underlinePattern("WAVY_DOUBLE")
                .underlineColorHex("FF0000") //red
                .build();

        TextValue textValue = new TextValue("formatted text", format);
        input.add("${formatted}", textValue);

        //get replacer instance
        DocxPlaceholderReplacer replacer = DocxPlaceholderReplacer.getInstance();
        //call service's replace method and pass input object
        byte[] resultDocx = replacer.replacePlaceholders(input);

        assertNotNull(resultDocx);
        assertTrue(resultDocx.length > 0);

//        File outputFile = new File("src/test/resources/result.docx");
//        Files.write(outputFile.toPath(), resultDocx);
    }
}