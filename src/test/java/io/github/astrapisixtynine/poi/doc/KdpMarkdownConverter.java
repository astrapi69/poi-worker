package io.github.astrapisixtynine.poi.doc;

import io.github.astrapi69.file.search.PathFinder;
import org.apache.poi.xwpf.usermodel.*;
import org.commonmark.node.*;
import org.commonmark.parser.Parser;
import org.commonmark.renderer.text.TextContentRenderer;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class KdpMarkdownConverter {

    private static final List<XWPFParagraph> tocEntries = new ArrayList<>();
    private static final Map<String, String> tocLinks = new HashMap<>();

    /**
     * Converts a Markdown file to a KDP-formatted DOCX with ToC and page breaks.
     *
     * @param mdFile   The input Markdown file
     * @param docxFile The output DOCX file
     * @throws IOException if file operations fail
     */
    public static void convertMdToKdpDocx(File mdFile, File docxFile) throws IOException {
        // Read Markdown content
        String markdown = readFile(mdFile);

        // Parse Markdown
        Parser parser = Parser.builder().build();
        Node document = parser.parse(markdown);

        // Create DOCX document
        try (XWPFDocument doc = new XWPFDocument();
             FileOutputStream out = new FileOutputStream(docxFile)) {

            // Insert a Table of Contents Placeholder
            XWPFParagraph tocTitle = doc.createParagraph();
            tocTitle.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun tocRun = tocTitle.createRun();
            tocRun.setBold(true);
            tocRun.setFontSize(18);
            tocRun.setText("Table of Contents");
            tocTitle.setPageBreak(true);

            // Process Markdown and generate content
            processMarkdownNode(doc, document);

            // Insert clickable Table of Contents
            insertTableOfContents(doc);

            // Save DOCX
            doc.write(out);
        }
        System.out.println("KDP-formatted DOCX with ToC created: " + docxFile.getAbsolutePath());
    }

    /**
     * Processes Markdown nodes and writes formatted content into the DOCX document.
     *
     * @param doc  The DOCX document
     * @param node The Markdown node
     */
    private static void processMarkdownNode(XWPFDocument doc, Node node) {
        node.accept(new AbstractVisitor() {
            @Override
            public void visit(Heading heading) {
                XWPFParagraph paragraph = doc.createParagraph();
                paragraph.setSpacingBefore(200);
                paragraph.setSpacingAfter(200);
                XWPFRun run = paragraph.createRun();
                run.setBold(true);
                run.setFontSize(getHeadingFontSize(heading.getLevel()));
                String headingText = getNodeText(heading);
                run.setText(headingText);

                // Add Bookmark for ToC
                String bookmarkId = "heading-" + tocEntries.size();
                tocLinks.put(headingText, bookmarkId);
                paragraph.setPageBreak(true);
                tocEntries.add(paragraph);
            }

            @Override
            public void visit(Paragraph paragraph) {
                XWPFParagraph docParagraph = doc.createParagraph();
                docParagraph.setSpacingAfter(150);
                XWPFRun run = docParagraph.createRun();
                run.setText(getNodeText(paragraph));
            }

            @Override
            public void visit(StrongEmphasis bold) {
                XWPFParagraph paragraph = doc.createParagraph();
                XWPFRun run = paragraph.createRun();
                run.setBold(true);
                run.setText(getNodeText(bold));
            }

            @Override
            public void visit(Emphasis italic) {
                XWPFParagraph paragraph = doc.createParagraph();
                XWPFRun run = paragraph.createRun();
                run.setItalic(true);
                run.setText(getNodeText(italic));
            }

            @Override
            public void visit(BulletList list) {
                Node listItem = list.getFirstChild();
                while (listItem != null) {
                    if (listItem instanceof ListItem) {
                        XWPFParagraph paragraph = doc.createParagraph();
                        paragraph.setIndentationLeft(500);
                        XWPFRun run = paragraph.createRun();
                        run.setText("• " + extractTextFromNode(listItem));
                    }
                    listItem = listItem.getNext();
                }
            }
        });
    }

    /**
     * Inserts a Table of Contents at the beginning of the document.
     *
     * @param doc The DOCX document
     */
    private static void insertTableOfContents(XWPFDocument doc) {
        if (tocEntries.isEmpty()) {
            return;
        }

        XWPFParagraph tocParagraph = doc.createParagraph();
        XWPFRun run = tocParagraph.createRun();
        run.setBold(true);
        run.setFontSize(14);
        run.setText("Table of Contents:");

        for (String heading : tocLinks.keySet()) {
            XWPFParagraph entryParagraph = doc.createParagraph();
            XWPFRun entryRun = entryParagraph.createRun();
            entryRun.setColor("0000FF"); // Kindle-friendly blue
            entryRun.setUnderline(UnderlinePatterns.SINGLE);
            entryRun.setText(heading);
        }
    }

    /**
     * Reads file content as a string.
     *
     * @param file The file to read
     * @return The file content
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

    /**
     * Extracts text from a Markdown node.
     *
     * @param node The Markdown node
     * @return The extracted text
     */
    private static String getNodeText(Node node) {
        return TextContentRenderer.builder().build().render(node);
    }

    /**
     * Extracts text from a node and its children.
     *
     * @param node The Markdown node
     * @return The extracted text
     */
    private static String extractTextFromNode(Node node) {
        StringBuilder text = new StringBuilder();
        Node child = node.getFirstChild();
        while (child != null) {
            text.append(getNodeText(child)).append(" ");
            child = child.getNext();
        }
        return text.toString().trim();
    }

    /**
     * Determines font size based on heading level.
     *
     * @param level The heading level (1-6)
     * @return The corresponding font size
     */
    private static int getHeadingFontSize(int level) {
        switch (level) {
            case 1: return 26;  // H1
            case 2: return 22;  // H2
            case 3: return 18;  // H3
            case 4: return 16;  // H4
            case 5: return 14;  // H5
            case 6: return 12;  // H6
            default: return 12;
        }
    }

    public static void main(String[] args) throws IOException {

        String docxFileName = "eternity-book.docx"; // Change this path
        String mdFileName = "eternity-book.md"; // Change this path
        File mdFile = new File(PathFinder.getSrcTestResourcesDir(), mdFileName);
        File docxFile = new File(PathFinder.getSrcTestResourcesDir(), docxFileName);

        convertMdToKdpDocx(mdFile, docxFile);
        convertMdToKdpDocx(mdFile, docxFile);
    }
}
