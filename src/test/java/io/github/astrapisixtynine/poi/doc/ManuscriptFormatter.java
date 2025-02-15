package io.github.astrapisixtynine.poi.doc;

import io.github.astrapi69.file.search.PathFinder;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;

public class ManuscriptFormatter {

    public static void main(String[] args) {
        String germanFileName = "Essaybuch.docx"; // Change this path
        String englishFileName = "eternity-book.docx"; // Change this path
        String outputFileName = "eternity-book-formatted.docx"; // Change this path


        File germanFile = new File(PathFinder.getSrcTestResourcesDir(), germanFileName);
        File englishFile = new File(PathFinder.getSrcTestResourcesDir(), englishFileName);
        File outputFile = new File(PathFinder.getSrcTestResourcesDir(), outputFileName);

        try {
            // Load the German manuscript as a formatting reference
            XWPFDocument germanDoc = new XWPFDocument(new FileInputStream(germanFile));
            
            // Load the English manuscript to apply formatting
            XWPFDocument englishDoc = new XWPFDocument(new FileInputStream(englishFile));

            // Create a new document for the formatted English version
            XWPFDocument formattedDoc = new XWPFDocument();

            // Apply formatting from the German document
            applyFormatting(germanDoc, englishDoc, formattedDoc);

            // Save the formatted document
            FileOutputStream out = new FileOutputStream(outputFile);
            formattedDoc.write(out);
            out.close();

            System.out.println("✅ Formatting complete! File saved to: " + outputFileName);

        } catch (IOException e) {
            System.err.println("❌ Error processing documents: " + e.getMessage());
        }
    }

    private static void applyFormatting(XWPFDocument germanDoc, XWPFDocument englishDoc, XWPFDocument formattedDoc) {
        // Loop through paragraphs in the English document
        for (XWPFParagraph englishPara : englishDoc.getParagraphs()) {
            // Create a new paragraph in the formatted document
            XWPFParagraph formattedPara = formattedDoc.createParagraph();

            // Try to match the formatting from the German document (same paragraph index)
            int index = englishDoc.getParagraphs().indexOf(englishPara);
            if (index < germanDoc.getParagraphs().size()) {
                XWPFParagraph germanPara = germanDoc.getParagraphs().get(index);
                formattedPara.setAlignment(germanPara.getAlignment());
                formattedPara.setSpacingAfter(germanPara.getSpacingAfter());
                formattedPara.setSpacingBefore(germanPara.getSpacingBefore());
                formattedPara.setStyle(germanPara.getStyle());
            }

            // Copy runs (text + formatting)
            for (XWPFRun englishRun : englishPara.getRuns()) {
                XWPFRun formattedRun = formattedPara.createRun();
                formattedRun.setText(englishRun.text());
                formattedRun.setBold(englishRun.isBold());
                formattedRun.setItalic(englishRun.isItalic());
                formattedRun.setFontSize(englishRun.getFontSize() > 0 ? englishRun.getFontSize() : 12);
                formattedRun.setFontFamily(englishRun.getFontFamily());
            }
        }
    }
}
