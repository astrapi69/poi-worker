package io.github.astrapisixtynine.poi.md;

import java.io.*;
import java.nio.file.*;
import java.util.regex.*;

import org.apache.poi.xwpf.usermodel.*;

import io.github.astrapi69.file.search.PathFinder;

public class MarkdownToDocxNoFlexmark
{
	public static void main(String[] args)
	{
		String inputPath; // Markdown file
		String outputPath; // Output DOCX


		String docxFileName = "eternity-book.docx";
		String mdFileName = "eternity-book.md";
		File mdFile = new File(PathFinder.getSrcTestResourcesDir(), mdFileName);
		File docxFile = new File(PathFinder.getSrcTestResourcesDir(), docxFileName);
		inputPath = mdFile.toPath().toString();
		outputPath = docxFile.toPath().toString();

		convert(inputPath, outputPath);
	}

	private static void convert(String inputPath, String outputPath)
	{
		try
		{
			// Read Markdown file
			String markdown = Files.readString(Paths.get(inputPath));

			// Create a new Word document
			XWPFDocument doc = new XWPFDocument();

			// Process Markdown manually
			parseMarkdown(markdown, doc);

			// Save DOCX
			try (FileOutputStream out = new FileOutputStream(outputPath))
			{
				doc.write(out);
			}

			System.out.println("DOCX created successfully at: " + outputPath);

		}
		catch (IOException e)
		{
			e.printStackTrace();
		}
	}

	private static void parseMarkdown(String markdown, XWPFDocument doc)
	{
		String[] lines = markdown.split("\n");

		for (String line : lines)
		{
			line = line.trim();

			// Headings
			if (line.startsWith("# "))
			{
				addHeading(doc, line.replace("# ", ""), 1);
			}
			else if (line.startsWith("## "))
			{
				addHeading(doc, line.replace("## ", ""), 2);
			}
			else if (line.startsWith("### "))
			{
				addHeading(doc, line.replace("### ", ""), 3);
			}
			// Bullet Points
			else if (line.startsWith("- ") || line.startsWith("* "))
			{
				addListItem(doc, line.substring(2));
			}
			// Numbered List
			else if (line.matches("^\\d+\\.\\s.*"))
			{
				addNumberedList(doc, line);
			}
			// Bold & Italic Formatting
			else if (!line.isEmpty())
			{
				addFormattedParagraph(doc, line);
			}
		}
	}

	private static void addHeading(XWPFDocument doc, String text, int level)
	{
		XWPFParagraph paragraph = doc.createParagraph();
		XWPFRun run = paragraph.createRun();
		paragraph.setStyle("Heading" + Math.min(level, 3)); // Limit to Heading 3
		run.setText(text);
		run.setBold(true);
	}

	private static void addListItem(XWPFDocument doc, String text)
	{
		XWPFParagraph paragraph = doc.createParagraph();
		paragraph.setIndentationLeft(500);
		XWPFRun run = paragraph.createRun();
		run.setText("• " + text);
	}

	private static void addNumberedList(XWPFDocument doc, String text)
	{
		XWPFParagraph paragraph = doc.createParagraph();
		paragraph.setIndentationLeft(500);
		XWPFRun run = paragraph.createRun();
		run.setText(text);
	}

	private static void addFormattedParagraph(XWPFDocument doc, String text)
	{
		XWPFParagraph paragraph = doc.createParagraph();
		XWPFRun run = paragraph.createRun();

		// Bold (**text**)
		Pattern boldPattern = Pattern.compile("\\*\\*(.*?)\\*\\*");
		Matcher boldMatcher = boldPattern.matcher(text);
		if (boldMatcher.find())
		{
			run.setBold(true);
			text = text.replaceAll("\\*\\*(.*?)\\*\\*", boldMatcher.group(1));
		}

		// Italic (*text*)
		Pattern italicPattern = Pattern.compile("\\*(.*?)\\*");
		Matcher italicMatcher = italicPattern.matcher(text);
		if (italicMatcher.find())
		{
			run.setItalic(true);
			text = text.replaceAll("\\*(.*?)\\*", italicMatcher.group(1));
		}

		run.setText(text);
	}
}
