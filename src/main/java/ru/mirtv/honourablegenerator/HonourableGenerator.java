/*
 *
 */
package ru.mirtv.honourablegenerator;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;
import java.util.Scanner;
import java.util.function.Consumer;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;

/**
 *
 * @author Babikov_PV
 */
public class HonourableGenerator {

  private static String templatePath = "c:\\tmp\\template1.pptx";
  private static String regardListPath = "c:\\tmp\\list.csv";
  private static String outputPath = "c:\\tmp\\output.pptx";

  /**
   * @param args the command line arguments
   */
  public static void main(String[] args) {
    XMLSlideShow template = null, output;
    XSLFTextRun fioTextRun = null, postTextRun = null, reasonTextRun = null;

    /*Scanner in = new Scanner(System.in);
    System.out.print("Введите полный путь до шаблона PowerPoint (.pptx): ");
    templatePath = in.next();
    System.out.print("Введите полный путь до файла со списком награждаемых (.csv): ");
    regardListPath = in.next();
    System.out.print("Куда сохранить результат (.pptx): ");
    outputPath = in.next();*/
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

      //get first slide from template and duplicate it on output file N times
      List<XSLFSlide> slides = template.getSlides();
      XSLFSlide mainSlide = slides.get(0);
      List<XSLFShape> shapes = mainSlide.getShapes();
      for (XSLFShape shape : shapes) {
        if (shape.getClass().toString().endsWith("XSLFTextBox")) {
          XSLFTextBox textBox = (XSLFTextBox) shape;
          List<XSLFTextParagraph> textParagraphs = textBox.getTextParagraphs();
          for (XSLFTextParagraph paragraph : textParagraphs) {
            List<XSLFTextRun> textRuns = paragraph.getTextRuns();
            for (XSLFTextRun textRun : textRuns) {
              String text = textRun.getRawText();
              if (text.contains("ФИО")) {
                System.out.println("found fio box");
                fioTextRun = textRun;
              } else if (text.contains("должность")) {
                System.out.println("found post box");
                postTextRun = textRun;
              } else if (text.contains("причина")) {
                System.out.println("found reason box");
                reasonTextRun = textRun;
              }
            }
          }
        }
      }
      
      if(fioTextRun != null && postTextRun != null && reasonTextRun != null){
        //read data from csv file
        try (Scanner csvReader = new Scanner(new File(regardListPath))) {
          csvReader.useDelimiter(";");
          while (csvReader.hasNext()) {
            fioTextRun.setText(csvReader.next().replace("\"","").trim());
            postTextRun.setText(csvReader.next().trim());
            reasonTextRun.setText(csvReader.next().replace("\"","").trim());
            output.createSlide().importContent(mainSlide);
          }
        } catch (FileNotFoundException ex) {
          System.err.println("Не удалось открыть файл со списком сотрудников.");
        }
      }

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
