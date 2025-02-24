package io.github.astrapisixtynine.poi.md.parse;

import static org.junit.jupiter.api.Assertions.*;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

/**
 * Parameterized unit test class for {@link MarkdownToDocxConverter} using CSV file data
 *
 * This class tests the formatted paragraph conversion using markdown input and expected output
 */
public class MarkdownToDocxConverterParameterizedTest
{
	/**
	 * Tests the conversion of formatted Markdown to DOCX
	 *
	 * @param markdownInput
	 *            The input Markdown text
	 * @param expectedText
	 *            The expected extracted text
	 * @param bold
	 *            Whether the text should be bold
	 * @param italic
	 *            Whether the text should be italic
	 */
	@ParameterizedTest
	@CsvSource({ "'**bold**', 'bold', true, false", "'*italic*', 'italic', false, true",
			"'**bold** and *italic*', 'bold and italic', true, true",
			"'***bold and italic***', 'bold and italic', true, true",
			"'___bold and italic___', 'bold and italic', true, true",
			"'** bold**', ' bold', false, false", "'**bold **', 'bold ', false, false",
			"'**bold *italic inside* bold**', 'bold italic inside bold', true, true",
			"'*italic **bold inside** italic*', 'italic bold inside italic', true, true",
			"'No formatting', 'No formatting', false, false",
			"'text**bold**text', 'textboldtext', true, false",
			"'text*italic*text', 'textitalictext', false, true",
			"'**bold & italic**', 'bold & italic', true, false",
			"'*italic with 123 numbers*', 'italic with 123 numbers', false, true",
			// "'\\*escaped*', '*escaped*', false, false",
			// "'\\_escaped_', '_escaped_', false, false",
			"'**bold *italic**', 'bold *italic', true, false" })
	void testFormattedParagraphConversion(String markdownInput, String expectedText, boolean bold,
		boolean italic)
	{
		XWPFDocument doc = new XWPFDocument();

		// Strictly match ***bold and italic*** or ___bold and italic___
		markdownInput = markdownInput.replaceAll("(\\*\\*\\*|___)(.*?)\\1", "$2");

		// Strictly match **bold** but not **bold* or bold**
		markdownInput = markdownInput.replaceAll("(?<!\\*)\\*\\*(?!\\*)(.*?)\\*\\*(?!\\*)", "$1");

		// Strictly match *italic* but not **italic* or italic**
		markdownInput = markdownInput.replaceAll("(?<!\\*)\\*(?!\\*)(.*?)\\*(?!\\*)", "$1");

		MarkdownToDocxConverter.addFormattedParagraph(doc, markdownInput, bold, italic);
		XWPFParagraph paragraph = doc.getParagraphs().get(0);
		XWPFRun run = paragraph.getRuns().get(0);

		assertEquals(expectedText, run.getText(0), "Extracted text should match expected text");
		assertEquals(bold, run.isBold(), "Bold flag mismatch");
		assertEquals(italic, run.isItalic(), "Italic flag mismatch");
	}
}
