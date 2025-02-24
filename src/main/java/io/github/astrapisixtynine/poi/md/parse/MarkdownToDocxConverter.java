package io.github.astrapisixtynine.poi.md.parse;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * Utility class for converting Markdown text into a DOCX document
 */
public class MarkdownToDocxConverter
{

	/**
	 * Converts a Markdown file into a DOCX document
	 *
	 * @param inputPath
	 *            the path to the Markdown file
	 * @param outputPath
	 *            the path where the DOCX file should be saved
	 */
	public static void convert(String inputPath, String outputPath) throws IOException
	{
		String markdown = Files.readString(Paths.get(inputPath));
		XWPFDocument doc = new XWPFDocument();
		parseMarkdown(markdown, doc);

		try (FileOutputStream out = new FileOutputStream(outputPath))
		{
			doc.write(out);
		}
	}

	/**
	 * Parses Markdown content and converts it into a DOCX document structure
	 *
	 * @param markdown
	 *            the Markdown content as a string
	 * @param doc
	 *            the DOCX document to populate
	 */
	public static void parseMarkdown(String markdown, XWPFDocument doc)
	{
		String[] lines = markdown.split("\\n");

		for (String line : lines)
		{
			line = line.trim();

			if (line.matches("#{1,6} .+"))
			{
				addHeading(doc, line.replaceAll("#+ ", ""), line.indexOf(' '));
			}
			else if (line.startsWith("- ") || line.startsWith("* "))
			{
				addListItem(doc, line.substring(2));
			}
			else if (line.matches("^\\d+\\.\\s.*"))
			{
				addNumberedList(doc, line);
			}
			else if (line.matches(".*\\*\\*\\*.*\\*\\*\\*.*"))
			{
				addFormattedParagraph(doc, line, true, true);
			}
			else if (line.matches(".*\\*\\*.*\\*\\*.*"))
			{
				addFormattedParagraph(doc, line, true, false);
			}
			else if (line.matches(".*\\*.*\\*.*"))
			{
				addFormattedParagraph(doc, line, false, true);
			}
			else if (line.startsWith("> "))
			{
				addBlockquote(doc, line.substring(2));
			}
			else if (line.matches("(---|\\*\\*\\*|___)"))
			{
				addHorizontalRule(doc);
			}
			else if (line.startsWith("```"))
			{
				addCodeBlock(doc, line);
			}
			else if (line.matches("!\\[.*\\]\\(.*\\)"))
			{
				addImagePlaceholder(doc, line);
			}
			else if (line.matches("\\[.*\\]\\(.*\\)"))
			{
				addLink(doc, line);
			}
			else if (!line.isEmpty())
			{
				addPlainParagraph(doc, line);
			}
		}
	}

	/**
	 * Adds a heading to the DOCX document
	 *
	 * @param doc
	 *            the DOCX document
	 * @param text
	 *            the heading text
	 * @param level
	 *            the heading level (1 to 6)
	 */
	public static void addHeading(XWPFDocument doc, String text, int level)
	{
		XWPFParagraph paragraph = doc.createParagraph();
		XWPFRun run = paragraph.createRun();
		paragraph.setStyle("Heading" + Math.min(level, 6));
		run.setText(text);
		run.setBold(true);
	}

	/**
	 * Adds an italicized paragraph to the DOCX document
	 *
	 * @param doc
	 *            the DOCX document
	 * @param text
	 *            the text to format as italic
	 */
	public static void addItalicParagraph(XWPFDocument doc, String text)
	{
		XWPFParagraph paragraph = doc.createParagraph();
		XWPFRun run = paragraph.createRun();
		run.setText(text.replaceAll("(?<!\\*)\\*(.*?)\\*(?!\\*)", "$1"));
		run.setItalic(true);
	}

	/**
	 * Adds a bullet list item to the DOCX document
	 *
	 * @param doc
	 *            the DOCX document
	 * @param text
	 *            the list item text
	 */
	public static void addListItem(XWPFDocument doc, String text)
	{
		XWPFParagraph paragraph = doc.createParagraph();
		paragraph.setIndentationLeft(500);
		XWPFRun run = paragraph.createRun();
		run.setText("• " + text);
	}

	/**
	 * Adds a numbered list item to the DOCX document
	 *
	 * @param doc
	 *            the DOCX document
	 * @param text
	 *            the numbered list text
	 */
	public static void addNumberedList(XWPFDocument doc, String text)
	{
		XWPFParagraph paragraph = doc.createParagraph();
		paragraph.setIndentationLeft(500);
		XWPFRun run = paragraph.createRun();
		run.setText(text);
	}

	/**
	 * Adds a blockquote to the DOCX document
	 *
	 * @param doc
	 *            the DOCX document
	 * @param text
	 *            the blockquote text
	 */
	public static void addBlockquote(XWPFDocument doc, String text)
	{
		XWPFParagraph paragraph = doc.createParagraph();
		paragraph.setIndentationLeft(800);
		XWPFRun run = paragraph.createRun();
		run.setText(text);
		run.setItalic(true);
	}

	/**
	 * Adds a horizontal rule to the DOCX document
	 *
	 * @param doc
	 *            the DOCX document
	 */
	public static void addHorizontalRule(XWPFDocument doc)
	{
		XWPFParagraph paragraph = doc.createParagraph();
		XWPFRun run = paragraph.createRun();
		run.setText("----------------------------");
	}

	/**
	 * Adds a code block to the DOCX document
	 *
	 * @param doc
	 *            the DOCX document
	 * @param code
	 *            the code block text
	 */
	public static void addCodeBlock(XWPFDocument doc, String code)
	{
		XWPFParagraph paragraph = doc.createParagraph();
		XWPFRun run = paragraph.createRun();
		run.setFontFamily("Courier New");
		run.setText(code.replace("```", ""));
	}

	/**
	 * Adds an image placeholder to the DOCX document
	 *
	 * @param doc
	 *            the DOCX document
	 * @param markdown
	 *            the Markdown image syntax
	 */
	public static void addImagePlaceholder(XWPFDocument doc, String markdown)
	{
		XWPFParagraph paragraph = doc.createParagraph();
		XWPFRun run = paragraph.createRun();
		run.setText("[Image Placeholder] " + markdown);
	}

	/**
	 * Adds a hyperlink to the DOCX document
	 *
	 * @param doc
	 *            the DOCX document
	 * @param markdown
	 *            the Markdown link syntax
	 */
	public static void addLink(XWPFDocument doc, String markdown)
	{
		String text = markdown.replaceAll("\\[(.*?)\\]\\((.*?)\\)", "$1 ($2)");
		XWPFParagraph paragraph = doc.createParagraph();
		XWPFRun run = paragraph.createRun();
		run.setText(text);
	}

	/**
	 * Adds a plain paragraph to the DOCX document
	 *
	 * @param doc
	 *            the DOCX document
	 * @param text
	 *            the paragraph text
	 */
	public static void addPlainParagraph(XWPFDocument doc, String text)
	{
		XWPFParagraph paragraph = doc.createParagraph();
		XWPFRun run = paragraph.createRun();
		run.setText(text);
	}

	/**
	 * Adds a formatted paragraph to the DOCX document with bold and italic styling
	 *
	 * @param doc
	 *            the DOCX document
	 * @param text
	 *            the paragraph text
	 * @param isBold
	 *            whether the text should be bold
	 * @param isItalic
	 *            whether the text should be italic
	 */
	public static void addFormattedParagraph(XWPFDocument doc, String text, boolean isBold,
		boolean isItalic)
	{
		XWPFParagraph paragraph = doc.createParagraph();
		XWPFRun run = paragraph.createRun();

		text = text.replaceAll("\\*\\*\\*(.*?)\\*\\*\\*", "$1"); // ***bold and italic***
		text = text.replaceAll("\\*\\*(.*?)\\*\\*", "$1"); // **bold**
		text = text.replaceAll("(?<!\\*)\\*(.*?)\\*(?!\\*)", "$1"); // *italic*

		run.setText(text);
		run.setBold(isBold);
		run.setItalic(isItalic);
	}
}
