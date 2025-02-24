package io.github.astrapisixtynine.poi.md.parse;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

/**
 * Unit tests for {@link MarkdownToDocxConverter}
 */
class MarkdownToDocxConverterTest
{
	private XWPFDocument doc;

	/**
	 * Sets up a new DOCX document before each test
	 */
	@BeforeEach
	void setUp()
	{
		doc = new XWPFDocument();
	}

	/**
	 * Tests that a Markdown file is successfully converted into a DOCX file
	 *
	 * @throws IOException
	 *             if an I/O error occurs
	 */
	@Test
	void testConvert() throws IOException
	{
		Path inputPath = Files.createTempFile("test-markdown", ".md");
		Path outputPath = Files.createTempFile("test-output", ".docx");
		Files.writeString(inputPath, "# Heading 1\nSome text\n");

		MarkdownToDocxConverter.convert(inputPath.toString(), outputPath.toString());

		assertTrue(Files.exists(outputPath), "Output DOCX should be created");
	}

	/**
	 * Tests that Markdown is correctly parsed into DOCX structure
	 */
	@Test
	void testParseMarkdown()
	{
		String markdown = "# Heading 1\nSome text\n- List item\n1. Numbered item\n";
		MarkdownToDocxConverter.parseMarkdown(markdown, doc);
		List<XWPFParagraph> paragraphs = doc.getParagraphs();
		assertEquals(4, paragraphs.size(), "Document should have correct number of paragraphs");
	}

	/**
	 * Tests adding an empty Markdown document
	 */
	@Test
	void testParseEmptyMarkdown()
	{
		String markdown = "";
		MarkdownToDocxConverter.parseMarkdown(markdown, doc);
		assertTrue(doc.getParagraphs().isEmpty(), "Document should be empty");
	}

	/**
	 * Tests adding a heading to the DOCX document
	 */
	@Test
	void testAddHeading()
	{
		MarkdownToDocxConverter.addHeading(doc, "Test Heading", 1);
		XWPFParagraph paragraph = doc.getParagraphs().get(0);
		assertEquals("Test Heading", paragraph.getText(), "Heading text should match");
	}

	/**
	 * Tests adding an italicized paragraph
	 */
	@Test
	void testAddItalicParagraph()
	{
		MarkdownToDocxConverter.addItalicParagraph(doc, "*italic text*");
		XWPFParagraph paragraph = doc.getParagraphs().get(0);
		assertEquals("italic text", paragraph.getText(), "Italic text should match");
	}

	/**
	 * Tests adding a bullet list item
	 */
	@Test
	void testAddListItem()
	{
		MarkdownToDocxConverter.addListItem(doc, "List item");
		XWPFParagraph paragraph = doc.getParagraphs().get(0);
		assertEquals("• List item", paragraph.getText(), "List item should be formatted correctly");
	}

	/**
	 * Tests adding a numbered list item
	 */
	@Test
	void testAddNumberedList()
	{
		MarkdownToDocxConverter.addNumberedList(doc, "1. Numbered item");
		XWPFParagraph paragraph = doc.getParagraphs().get(0);
		assertEquals("1. Numbered item", paragraph.getText(),
			"Numbered item should be formatted correctly");
	}

	/**
	 * Tests adding a blockquote
	 */
	@Test
	void testAddBlockquote()
	{
		MarkdownToDocxConverter.addBlockquote(doc, "Blockquote text");
		XWPFParagraph paragraph = doc.getParagraphs().get(0);
		assertEquals("Blockquote text", paragraph.getText(), "Blockquote should match");
	}

	/**
	 * Tests adding a horizontal rule
	 */
	@Test
	void testAddHorizontalRule()
	{
		MarkdownToDocxConverter.addHorizontalRule(doc);
		XWPFParagraph paragraph = doc.getParagraphs().get(0);
		assertEquals("----------------------------", paragraph.getText(),
			"Horizontal rule should match");
	}

	/**
	 * Tests adding a code block
	 */
	@Test
	void testAddCodeBlock()
	{
		MarkdownToDocxConverter.addCodeBlock(doc, "```Code block```\n");
		XWPFParagraph paragraph = doc.getParagraphs().get(0);
		assertEquals("Code block", paragraph.getText().trim(), "Code block should match");
	}

	/**
	 * Tests adding an image placeholder
	 */
	@Test
	void testAddImagePlaceholder()
	{
		MarkdownToDocxConverter.addImagePlaceholder(doc, "![Alt text](image.jpg)");
		XWPFParagraph paragraph = doc.getParagraphs().get(0);
		assertEquals("[Image Placeholder] ![Alt text](image.jpg)", paragraph.getText(),
			"Image placeholder should match");
	}

	/**
	 * Tests adding a hyperlink
	 */
	@Test
	void testAddLink()
	{
		MarkdownToDocxConverter.addLink(doc, "[Link](http://example.com)");
		XWPFParagraph paragraph = doc.getParagraphs().get(0);
		assertEquals("Link (http://example.com)", paragraph.getText(),
			"Link should be formatted correctly");
	}

	/**
	 * Tests adding a plain paragraph
	 */
	@Test
	void testAddPlainParagraph()
	{
		MarkdownToDocxConverter.addPlainParagraph(doc, "Plain text");
		XWPFParagraph paragraph = doc.getParagraphs().get(0);
		assertEquals("Plain text", paragraph.getText(), "Plain paragraph should match");
	}
}
