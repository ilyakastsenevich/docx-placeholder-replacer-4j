package io.github.ilyakastsenevich.docxplaceholderreplacer4j.replacer;

import io.github.ilyakastsenevich.docxplaceholderreplacer4j.dto.ReplacePlaceholdersInput;

/**
 * A DocxPlaceholderReplacer is an interface that defines a method to replace placeholders in a DOCX file with given values.
 */
public interface DocxPlaceholderReplacer {

    /**
     * Replaces placeholders in a DOCX file with given values and returns the modified file as a byte array.
     * @param input a ReplacePlaceholdersInput object that contains the DOCX file and the values to replace the placeholders with
     * @return a byte array representing the modified DOCX file
     */
    byte[] replacePlaceholders(ReplacePlaceholdersInput input);

    /**
     * Returns an instance of DocxPlaceholderReplacer implementation.
     * @return a DocxPlaceholderReplacer object
     */
    static DocxPlaceholderReplacer getInstance() {
        return new DocxPlaceholderReplacerImpl();
    }

}
