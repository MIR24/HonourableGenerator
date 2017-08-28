/* This program allow to merge two files - .pptx (PowerPoint) template and
 * .xslx list of values. Program will replace custom placeholder in template  
 * and generate new file filled with the values from list.
 */
package ru.mirtv.honourablegenerator;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Scanner;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import ru.mirtv.honourable.objects.Regard;

/**
 *
 * @author Babikov_PV
 */
public class HonourableGenerator {

  private static String templatePath;
  private static String regardListPath;
  private static String outputPath;

  /**
   * @param args the command line arguments
   */
  public static void main(String[] args) {
    XMLSlideShow template = null, output;
    ArrayList<Regard> regards = new ArrayList<>();
    FileInputStream fis;
    XSSFWorkbook regardsList = null;

    Scanner in = new Scanner(System.in);
    System.out.print("Введите полный путь до шаблона PowerPoint (.pptx): ");
    templatePath = in.next();
    System.out.print("Введите полный путь до файла со списком награждаемых (.xlsx): ");
    regardListPath = in.next();
    System.out.print("Куда сохранить результат (.pptx): ");
    outputPath = in.next();

    //open template file
    try (FileInputStream is = new FileInputStream(new File(templatePath))) {
      template = new XMLSlideShow(is);
    } catch (FileNotFoundException ex) {
      System.err.println("Файл шаблона отсутствует или неверно указан путь.");
    } catch (IOException ex) {
      System.err.println("Не удалось прочесть файл шаблона.");
    }

    if (template != null) {
      //create output file
      output = new XMLSlideShow();
      output.setPageSize(template.getPageSize());

      //open xlsx file
      try {
        File xlsxFile = new File(regardListPath);
        fis = new FileInputStream(xlsxFile);
        regardsList = new XSSFWorkbook(fis);
      } catch (FileNotFoundException ex) {
        System.err.println("Файл со списком награждаемых не найден.");
      } catch (IOException ex) {
        System.err.println("Не удалось прочитать список награждаемых.");
      }

      System.out.println("Считываем данные из списка.");
      //read all data from xlsx to array
      if (regardsList != null) {
        XSSFSheet mainSheet = regardsList.getSheetAt(0);
        Iterator<Row> rowIterator = mainSheet.iterator();
        while (rowIterator.hasNext()) {
          Row row = rowIterator.next();
          if (row.getCell(0) != null) {
            String name = row.getCell(0).getStringCellValue();
            String post = row.getCell(1).getStringCellValue();
            String reason = row.getCell(2).getStringCellValue();
            regards.add(new Regard(name, post, reason));
          }
        }
      }

      if (regards.size() == 0) {
        System.err.println("Список был пуст или некорректен.");
      } else {
        //get first slide of template - it most not be changed
        List<XSLFSlide> slides = template.getSlides();
        XSLFSlide mainSlide = slides.get(0);

        System.out.println("Заполняем слайды.");
        //iterate over all regards in list and add new slide for each
        regards.forEach((Regard regard) -> {
          XSLFSlide slide = output.createSlide().importContent(mainSlide);
          List<XSLFShape> shapes = slide.getShapes();
          for (XSLFShape shape : shapes) {
            //check if shape is instance of XSLFTextBox
            if (shape.getClass().toString().endsWith("XSLFTextBox")) {
              XSLFTextBox textBox = (XSLFTextBox) shape;
              List<XSLFTextParagraph> textParagraphs = textBox.getTextParagraphs();
              for (XSLFTextParagraph paragraph : textParagraphs) {
                List<XSLFTextRun> textRuns = paragraph.getTextRuns();
                textRuns.forEach((textRun) -> {
                  String text = textRun.getRawText();
                  //this is a keyword in template file that must be replaced
                  if (text.contains("ФИО")) {
                    textRun.setText(regard.getName());
                  } else if (text.contains("должность")) {
                    textRun.setText(regard.getPost());
                  } else if (text.contains("причина")) {
                    textRun.setText(regard.getReason());
                  }
                });
              }
            }
          }
        });

        System.out.println("Сохраняем результат.");
        //save output
        try {
          File outputFile = new File(outputPath);
          FileOutputStream out = new FileOutputStream(outputFile);
          output.write(out);
        } catch (FileNotFoundException ex) {
          System.err.println("Результирующий файл не создан. Проверьте путь.");
        } catch (IOException ex) {
          System.err.println("Не удалось записать результирующий файл.");
        }
      }

      //close all files
      try {
        template.close();
        output.close();
      } catch (IOException ex) {
        System.err.println("Ошибка при закрытии файла.");
      }
    }
  }

}
