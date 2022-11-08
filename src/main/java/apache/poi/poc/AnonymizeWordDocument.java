package apache.poi.poc;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.List;
import java.util.Set;
import java.util.stream.Collectors;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class AnonymizeWordDocument {

  public static void main(String[] args) throws IOException {
// .docx
//   new AnonymizeWordDocument().updateMsWordDocxFile("c:\\work\\Abiola_O_Obatuase_resume.docx", "c:\\work\\Abiola_O_Obatuase_resume_upd.docx", "4109985643,443722987656,nurse1@gmail.com");

//    new AnonymizeWordDocument().updateMsWordDocxFile("c:\\work\\JASVIR_Sun.docx", "c:\\work\\JASVIR_Sun_upd.docx", "(703) 502-8672,sunnyside@GMAIL.COM");
//    new AnonymizeWordDocument().updateMsWordDocxFile("c:\\work\\Adenike_Akonnursing_resume.docx", "c:\\work\\Adenike_Akonnursing_resume_upd.docx", "240-305-0100,anurse@gmail.com");
//    new AnonymizeWordDocument().updateMsWordDocxFile("c:\\work\\SenthilKumar.docx", "c:\\work\\SenthilKumar_upd.docx","3176029380,senthilkumar.lakshma@gmail.com");
//
//    // .doc
    new AnonymizeWordDocument().updateMsWordDocxFile("resumes\\cv_en.docx", "resumes\\cv_en_upd.docx","Myllon.tennyson@info.co.in,0021-1258-32698,0054-45897-4562");
    new AnonymizeWordDocument().updateMsWordDocFile("resumes\\Doctor_Resume_MsWord_2MB.doc", "resumes\\Doctor_Resume_MsWord_2MB_upd.doc","my_is@yahoo.com,010-10000001");
    new AnonymizeWordDocument().updateMsWordDocxFile("resumes\\n_test.docx", "resumes\\n_test_1_upd.docx", "cbooth@yahoo.com,cboh@gmail.com,cbooth@volny.cz,3176029380,Test@gmail.com");
    new AnonymizeWordDocument().updateMsWordDocxFile("resumes\\Blair_Brewer_Resume (1).docx", "resumes\\Blair_Brewer_Resume (1)_upd.docx", "(434) 466-6622,brew@gmail.com");
    new AnonymizeWordDocument().updateMsWordDocFile("resumes\\cathy_resume.doc", "resumes\\cathy_resume_upd.doc","cbooth@yahoo.com,(301) 213-7771");
    new AnonymizeWordDocument().updateMsWordDocFile("resumes\\Demo_Doc.doc", "resumes\\Demo_Doc_upd.doc","harriet.smith.doc@hotmail.co.uk,07825498149");
    new AnonymizeWordDocument().updateMsWordDocFile("resumes\\Tim Mccall__09_24_2015.doc", "resumes\\Tim Mccall__09_24_2015_upd.doc","Jnurse7@yahoo.com,443-929-8888");
    new AnonymizeWordDocument().updateMsWordDocFile("resumes\\Doctor_Resume_MsWord_2.2MB.doc", "resumes\\Doctor_Resume_MsWord_2.2MB_upd.doc","my_is@yahoo.com,010-10000001");
  }

  private void updateMsWordDocxFile(String input, String output, String words) throws IOException {

      var document = new XWPFDocument(Files.newInputStream(Paths.get(input)));

    Set<String> wordSet = Arrays.stream(words.split(",")).collect(Collectors.toSet());
    processParagraphs(wordSet, document.getParagraphs());
    processTables(wordSet, document.getTables());
    processHeaders(wordSet, document.getHeaderList());
    processFooters(wordSet, document.getFooterList());
    document.write(new FileOutputStream(output));
    document.close();
  }

  private void processParagraphs(Set<String> wordSet, List<XWPFParagraph> paragraphs) {
    for (XWPFParagraph paragraph : paragraphs) {
      var paragraphText = paragraph.getText();
      if (StringUtils.isNotEmpty(paragraphText)) {
        wordSet.forEach(word -> {
          if (paragraphText.contains(word)) {
            for (XWPFRun xwpfRun : paragraph.getRuns()) {
              String text = xwpfRun.getText(0);
              if (StringUtils.isNotEmpty(text)) {
                processRunText(word, xwpfRun, text);
              }
            }
          }
        });
      }
    }
  }

  private void processRunText(String word, XWPFRun xwpfRun, String text) {
    var runLength = text.length();
    if (word.equals(text)) {
      xwpfRun.setText(StringUtils.repeat("*", runLength), 0);
    } else if (text.contains(word)) {
      xwpfRun.setText(text.replace(word, StringUtils.repeat("*", word.length())), 0);
    } else if (word.contains(text.trim())) {
      xwpfRun.setText(StringUtils.repeat("*", runLength), 0);
    }
  }

  private void processHeaders(Set<String> wordSet, List<XWPFHeader> headers) {
    for (XWPFHeader header : headers) {
      if (StringUtils.isNotEmpty(header.getText())) {
        processParagraphs(wordSet, header.getParagraphs());
        processTables(wordSet, header.getTables());
      }
    }
  }

  private void processFooters(Set<String> wordSet, List<XWPFFooter> footers) {
    for (XWPFFooter footer : footers) {
      if (StringUtils.isNotEmpty(footer.getText())) {
        processParagraphs(wordSet, footer.getParagraphs());
        processTables(wordSet, footer.getTables());
      }
    }
  }

  private void processTables(Set<String> wordSet, List<XWPFTable> tables) {
    for (XWPFTable table : tables) {
      if (StringUtils.isNotEmpty(table.getText())) {
        for (XWPFTableRow row : table.getRows()) {
          for (XWPFTableCell cell : row.getTableCells()) {
            processParagraphs(wordSet, cell.getParagraphs());
          }
        }
      }
    }
  }

  private void updateMsWordDocFile(String input, String output, String words) throws IOException {
    var document = new HWPFDocument(Files.newInputStream(Paths.get(input)));
    Set<String> wordSet = Arrays.stream(words.split(",")).collect(Collectors.toSet());

    Range range = document.getRange();
    wordSet.forEach(word -> range.replaceText(word, StringUtils.repeat("*", word.length())));
    document.write(new FileOutputStream(output));
    document.close();
  }
}
