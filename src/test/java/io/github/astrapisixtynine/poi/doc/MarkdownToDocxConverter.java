package io.github.astrapisixtynine.poi.doc;

import com.vladsch.flexmark.html.HtmlRenderer;
import com.vladsch.flexmark.parser.Parser;
import io.github.astrapi69.file.search.PathFinder;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import com.vladsch.flexmark.util.ast.Document;

import java.io.*;

public class MarkdownToDocxConverter {

    /**
     * Converts a Markdown file to a DOCX file.
     *
     * @param mdFile  the input Markdown file
     * @param docxFile the output DOCX file
     * @throws IOException if file operations fail
     */
    public static void convertMdToDocx(File mdFile, File docxFile) throws IOException {
        // Read Markdown content
        String markdown = readFile(mdFile);

        // Convert Markdown to HTML
        Parser parser = Parser.builder().build();
        Document document = parser.parse(markdown);
        HtmlRenderer renderer = HtmlRenderer.builder().build();
        String htmlContent = renderer.render(document);

        // Create a DOCX document
        try (XWPFDocument doc = new XWPFDocument();
             FileOutputStream out = new FileOutputStream(docxFile)) {

            XWPFParagraph paragraph = doc.createParagraph();
            XWPFRun run = paragraph.createRun();
            run.setText(htmlContent);

            // Save DOCX file
            doc.write(out);
        }
        System.out.println("Conversion successful: " + docxFile.getAbsolutePath());
    }

    /**
     * Reads file content as a string.
     *
     * @param file the file to read
     * @return the file content
     * @throws IOException if file reading fails
     */
    private static String readFile(File file) throws IOException {
        try (BufferedReader br = new BufferedReader(new FileReader(file))) {
            StringBuilder content = new StringBuilder();
            String line;
            while ((line = br.readLine()) != null) {
                content.append(line).append("\n");
            }
            return content.toString();
        }
    }

    public static void main(String[] args) throws IOException {

        String docxFileName = "eternity-book.docx"; // Change this path
        String mdFileName = "eternity-book.md"; // Change this path
        File mdFile = new File(PathFinder.getSrcTestResourcesDir(), mdFileName);
        File docxFile = new File(PathFinder.getSrcTestResourcesDir(), docxFileName);

        convertMdToDocx(mdFile, docxFile);
    }
}
