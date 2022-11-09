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
//   new AnonymizeWordDocument().updateMsWordDocxFile("c:\\work\\Abiola_O_Obatuase_resume.docx", "c:\\work\\Abiola_O_Obatuase_resume_upd.docx", "email_address,phone_number");

//    new AnonymizeWordDocument().updateMsWordDocxFile("c:\\work\\JASVIR_Sun.docx", "c:\\work\\JASVIR_Sun_upd.docx", "email_address,phone_number");
//    new AnonymizeWordDocument().updateMsWordDocxFile("c:\\work\\Adenike_Akonnursing_resume.docx", "c:\\work\\Adenike_Akonnursing_resume_upd.docx", "email_address,phone_number");
//    new AnonymizeWordDocument().updateMsWordDocxFile("c:\\work\\SenthilKumar.docx", "c:\\work\\SenthilKumar_upd.docx","email_address,phone_number");
//
//    // .doc
//    new AnonymizeWordDocument().updateMsWordDocxFile("resumes\\cv_en.docx", "resumes\\cv_en_upd.docx","email_address,phone_number,phone_number");
//    new AnonymizeWordDocument().updateMsWordDocFile("resumes\\Doctor_Resume_MsWord_2MB.doc", "resumes\\Doctor_Resume_MsWord_2MB_upd.doc","email_address,phone_number");
    new AnonymizeWordDocument().updateMsWordDocxFile("resumes\\n_test.docx", "resumes\\n_test_upd.docx", "cobondeno@yahoo.com,cobondeno@gmail.com,cobondeno@volny.cz,3176029380");
//    new AnonymizeWordDocument().updateMsWordDocxFile("resumes\\Blair_Brewer_Resume (1).docx", "resumes\\Blair_Brewer_Resume (1)_upd.docx", "email_address,phone_number");
//    new AnonymizeWordDocument().updateMsWordDocFile("resumes\\cathy_resume.doc", "resumes\\cathy_resume_upd.doc","email_address,phone_number");
//    new AnonymizeWordDocument().updateMsWordDocFile("resumes\\Demo_Doc.doc", "resumes\\Demo_Doc_upd.doc","email_address,phone_number");
//    new AnonymizeWordDocument().updateMsWordDocFile("resumes\\Tim Mccall__09_24_2015.doc", "resumes\\Tim Mccall__09_24_2015_upd.doc","email_address,phone_number");
//    new AnonymizeWordDocument().updateMsWordDocFile("resumes\\Doctor_Resume_MsWord_2.2MB.doc", "resumes\\Doctor_Resume_MsWord_2.2MB_upd.doc","email_address,phone_number");
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
      processParagraph(wordSet, paragraph);
    }
  }

  private void processParagraph(Set<String> wordSet, XWPFParagraph paragraph) {
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
