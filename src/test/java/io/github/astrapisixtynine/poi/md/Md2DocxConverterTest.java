package io.github.astrapisixtynine.poi.md;

import static org.junit.jupiter.api.Assertions.*;

import java.io.*;
import java.nio.file.Files;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

/**
 * Unit test for {@link Md2DocxConverter}
 */
class Md2DocxConverterTest
{

	private File mdFile;
	private File docxFile;

	@BeforeEach
	void setUp() throws IOException
	{
		mdFile = Files.createTempFile("test-markdown", ".md").toFile();
		docxFile = Files.createTempFile("test-output", ".docx").toFile();
	}

	@AfterEach
	void tearDown()
	{
		mdFile.delete();
		docxFile.delete();
	}

	@Test
	void testConvertMdToFormattedDocx() throws IOException
	{
		String markdownContent = "# Heading 1\n\nThis is a paragraph.\n\n**Bold Text**\n\n*Italic Text*\n\n- Item 1\n- Item 2";
		Files.write(mdFile.toPath(), markdownContent.getBytes());

		Md2DocxConverter.convertMdToFormattedDocx(mdFile, docxFile);

		assertTrue(docxFile.exists(), "DOCX file should be created");
		assertTrue(docxFile.length() > 0, "DOCX file should not be empty");

		try (XWPFDocument document = new XWPFDocument(new FileInputStream(docxFile)))
		{
			assertNotNull(document, "DOCX document should be readable");
			assertFalse(document.getParagraphs().isEmpty(), "Document should contain paragraphs");
			assertEquals(7, document.getParagraphs().size(), "Expected 7 paragraphs");
		}
	}

	@Test
	void testEmptyMarkdownFile() throws IOException
	{
		Files.write(mdFile.toPath(), new byte[0]);

		Md2DocxConverter.convertMdToFormattedDocx(mdFile, docxFile);

		assertTrue(docxFile.exists(), "DOCX file should be created");

		try (XWPFDocument document = new XWPFDocument(new FileInputStream(docxFile)))
		{
			assertNotNull(document, "DOCX document should be readable");
			assertTrue(document.getParagraphs().isEmpty(), "Document should have no paragraphs");
		}
	}
}
