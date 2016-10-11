import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import java.io.File;
import java.io.FileOutputStream;

public class WordTest {
    public static void main(String... args) throws Exception {
        try (FileOutputStream out = new FileOutputStream(new File("word_test.docx"))) {
            XWPFDocument document = new XWPFDocument();
            XWPFParagraph paragraph = document.createParagraph();
            XWPFRun run = paragraph.createRun();
            run.setText("サンプル");
            document.write(out);
        }
    }
}
