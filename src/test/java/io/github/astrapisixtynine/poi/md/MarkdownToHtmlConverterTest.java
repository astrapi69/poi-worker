package io.github.astrapisixtynine.poi.md;

import static org.assertj.core.api.Assertions.assertThat;

import java.io.File;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;

import org.junit.jupiter.api.Test;

import io.github.astrapi69.file.search.PathFinder;
import io.github.astrapi69.file.write.StoreFileExtensions;

class MarkdownToHtmlConverterTest
{

	@Test
	void testHeadingConversion()
	{
		String markdown = "# Heading 1\n## Heading 2\n### Heading 3";
		String html = MarkdownToHtmlConverter.convertMarkdownToHtml(markdown);

		assertThat(html).contains("<h1>Heading 1</h1>");
		assertThat(html).contains("<h2>Heading 2</h2>");
		assertThat(html).contains("<h3>Heading 3</h3>");
	}

	@Test
	void testMarkdownConversion() throws IOException
	{

		String mdFilePath;

		String htmlFileName = "eternity-book.html"; // Change this path
		String mdFileName = "eternity-book.md"; // Change this path
		File mdFile = new File(PathFinder.getSrcTestResourcesDir(), mdFileName);
		File htmlFile = new File(PathFinder.getSrcTestResourcesDir(), htmlFileName);

		mdFilePath = mdFile.toPath().toString();
		// Markdown-Datei lesen
		String markdownContent = new String(Files.readAllBytes(Paths.get(mdFilePath)),
			StandardCharsets.UTF_8);

		String htmlContent = MarkdownToHtmlConverter.convertMarkdownToHtml(markdownContent);
		StoreFileExtensions.toFile(htmlFile, htmlContent);
		htmlFile.delete();
	}

	@Test
	void testParagraphConversion()
	{
		String markdown = "This is a simple paragraph.";
		String html = MarkdownToHtmlConverter.convertMarkdownToHtml(markdown);

		assertThat(html).contains("<p>This is a simple paragraph.</p>");
	}

	@Test
	void testBoldAndItalicTextConversion()
	{
		String markdown = "**Bold Text**\n*Italic Text*";
		String html = MarkdownToHtmlConverter.convertMarkdownToHtml(markdown);

		assertThat(html).contains("<strong>Bold Text</strong>");
		assertThat(html).contains("<em>Italic Text</em>");
	}

	@Test
	void testUnorderedListConversion()
	{
		String markdown = "- Item 1\n- Item 2";
		String html = MarkdownToHtmlConverter.convertMarkdownToHtml(markdown);

		assertThat(html).contains("<ul>");
		assertThat(html).contains("<li>Item 1</li>");
		assertThat(html).contains("<li>Item 2</li>");
		assertThat(html).contains("</ul>");
	}

	@Test
	void testOrderedListConversion()
	{
		String markdown = "1. First Item\n2. Second Item";
		String html = MarkdownToHtmlConverter.convertMarkdownToHtml(markdown);

		assertThat(html).contains("<ol>");
		assertThat(html).contains("<li>First Item</li>");
		assertThat(html).contains("<li>Second Item</li>");
		assertThat(html).contains("</ol>");
	}

	@Test
	void testMixedContent()
	{
		String markdown = "# Title\nThis is **bold** text.\n\n- List Item 1\n- List Item 2";
		String html = MarkdownToHtmlConverter.convertMarkdownToHtml(markdown);

		assertThat(html).contains("<h1>Title</h1>");
		assertThat(html).contains("<p>This is <strong>bold</strong> text.</p>");
		assertThat(html).contains("<ul>");
		assertThat(html).contains("<li>List Item 1</li>");
		assertThat(html).contains("<li>List Item 2</li>");
		assertThat(html).contains("</ul>");
	}

	@Test
	void testEmptyMarkdown()
	{
		String markdown = "";
		String html = MarkdownToHtmlConverter.convertMarkdownToHtml(markdown);

		assertThat(html).isEqualTo("");
	}
}
