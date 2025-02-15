package io.github.astrapisixtynine.poi.doc;

import io.github.astrapi69.file.search.PathFinder;
import org.commonmark.node.*;
        import org.commonmark.parser.Parser;
import org.commonmark.renderer.text.TextContentRenderer;
import org.apache.poi.xwpf.usermodel.*;

        import java.io.*;

public class Md2Docx {

    /**
     * Converts a Markdown file to a DOCX file.
     *
     * @param mdFile   the input Markdown file
     * @param docxFile the output DOCX file
     * @throws IOException if file operations fail
     */
    public static void convertMdToDocx(File mdFile, File docxFile) throws IOException {
        // Read Markdown content
        String markdown = readFile(mdFile);

        // Convert Markdown to text
        Parser parser = Parser.builder().build();
        Node node = parser.parse(markdown);  // Corrected to Node
        TextContentRenderer renderer = TextContentRenderer.builder().build();
        String plainText = renderer.render(node);

        // Create a DOCX document
        try (XWPFDocument doc = new XWPFDocument();
             FileOutputStream out = new FileOutputStream(docxFile)) {

            XWPFParagraph paragraph = doc.createParagraph();
            XWPFRun run = paragraph.createRun();
            run.setText(plainText);
            run.setFontSize(12);

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

