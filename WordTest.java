import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.util.Units;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileInputStream;

public class WordTest {
    public static void main(String... args) throws Exception {
        try (FileOutputStream out = new FileOutputStream(new File("word_test.docx"))) {
            XWPFDocument document = new XWPFDocument();
            drawTitle(document);
            drawContent1(document);
            drawContent2(document);
            drawContent3(document);
            drawContent4(document);
            pageBreak(document);
            drawContent1(document);
            document.write(out);
        }
    }
    private static void drawTitle(XWPFDocument document) {
            XWPFParagraph paragraph = document.createParagraph();
            XWPFRun run = paragraph.createRun();
            run.setFontSize(20);
            run.setText("TITLE (font size is 20)");
    }
    private static void drawContent1(XWPFDocument document) {
            XWPFParagraph paragraph = document.createParagraph();
            XWPFRun run = paragraph.createRun();
            run.setText("text text text");
    }
    private static void drawContent2(XWPFDocument document) {
            XWPFParagraph paragraph = document.createParagraph();
            XWPFRun run = paragraph.createRun();
            run.setColor("FF0000");
            run.setText("Red text");
    }
    private static void drawContent3(XWPFDocument document) {
            XWPFParagraph paragraph = document.createParagraph();
            paragraph.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun run = paragraph.createRun();
            run.setText("Text align center");
    }
    private static void drawContent4(XWPFDocument document) throws Exception {
            XWPFParagraph paragraph = document.createParagraph();
            XWPFRun run = paragraph.createRun();
            String imgFile = "sample.png";
         try (FileInputStream image = new FileInputStream(imgFile)) {
            run.addPicture(image, XWPFDocument.PICTURE_TYPE_JPEG, imgFile, Units.toEMU(200), Units.toEMU(200)); // 200x200 pixels
         }
    }
    private static void pageBreak(XWPFDocument document) {
            XWPFParagraph paragraph = document.createParagraph();
            paragraph.setPageBreak(true);
    }
}
