package io.github.astrapisixtynine.poi.md;

import java.util.Arrays;

import org.jsoup.Jsoup;

import com.vladsch.flexmark.ext.emoji.EmojiExtension;
import com.vladsch.flexmark.ext.typographic.TypographicExtension;
import com.vladsch.flexmark.html.HtmlRenderer;
import com.vladsch.flexmark.parser.Parser;
import com.vladsch.flexmark.util.ast.Node;
import com.vladsch.flexmark.util.data.MutableDataSet;


public final class MarkdownToHtmlConverter
{

	/**
	 * Converts Markdown to HTML with support for emojis and smart punctuation.
	 *
	 * @param markdownContent
	 *            The Markdown text as a String
	 * @return Cleaned HTML as a String with emoji and punctuation fixes
	 */
	public static String convertMarkdownToHtml(String markdownContent)
	{
		// Configure Flexmark with emoji and smart punctuation support
		MutableDataSet options = new MutableDataSet();
		options.set(Parser.EXTENSIONS,
			Arrays.asList(EmojiExtension.create(), TypographicExtension.create()));

		Parser parser = Parser.builder(options).build();
		HtmlRenderer renderer = HtmlRenderer.builder(options).build();

		// Convert Markdown to HTML
		Node documentNode = parser.parse(markdownContent);
		String htmlContent = renderer.render(documentNode);

		// Clean HTML with JSoup
		org.jsoup.nodes.Document jsoupDocument = Jsoup.parse(htmlContent);
		return jsoupDocument.body().html(); // Return only body content
	}

	public static void main(String[] args)
	{
		String markdown = "Hello, world! :smile: :rocket:\n\nThis is a test -- with dashes and ellipses...\n\nAnother test --- em dash!";
		String html = convertMarkdownToHtml(markdown);
		System.out.println(html);
	}
}
