package com.example.filereader;

import com.example.filereader.util.ReaderDoc;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class ReadWhiteInDocTest {

    @Test
    public void readDocTheOneForm() throws IOException {

        ReaderDoc readerDoc = new ReaderDoc();
        readerDoc.read(ReaderTest.class.getResource("/example.docx").getPath());
    }

    @Test
    public void readDocTheOtherForm() throws IOException {

        try {

            /**
             * if uploaded doc then use HWPF else if uploaded Docx file use
             * XWPFDocument
             */
            XWPFDocument doc = new XWPFDocument(
                    OPCPackage.open(ReaderTest.class.getResource("/example.docx").getPath()));
            for (XWPFParagraph p : doc.getParagraphs()) {
                List<XWPFRun> runs = p.getRuns();
                if (runs != null) {
                    for (XWPFRun r : runs) {
                        String text = r.getText(0);
                        if (text != null && text.contains("$$key$$")) {
                            text = text.replace("$$key$$", "ABCD");//your content
                            r.setText(text, 0);
                        }
                    }
                }
            }

            for (XWPFTable tbl : doc.getTables()) {
                for (XWPFTableRow row : tbl.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph p : cell.getParagraphs()) {
                            for (XWPFRun r : p.getRuns()) {
                                String text = r.getText(0);
                                if (text != null && text.contains("$$key$$")) {
                                    text = text.replace("$$key$$", "abcd");
                                    r.setText(text, 0);
                                }
                            }
                        }
                    }
                }
            }

            doc.write(new FileOutputStream(ReaderTest.class.getResource("/exampleOut.docx").getPath()));
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } finally {

        }
    }
}
