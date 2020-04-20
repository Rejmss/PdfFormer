import com.itextpdf.kernel.color.DeviceRgb;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.*;
import com.itextpdf.kernel.pdf.canvas.PdfCanvas;
import com.itextpdf.kernel.pdf.xobject.PdfFormXObject;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Text;
import com.itextpdf.layout.property.TextAlignment;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

import java.io.*;

public class Main {
    private static FileInputStream file;
    private static HSSFWorkbook workbook;
    private static HSSFSheet sheet;
    private static String dstFolder = "dst/";
    private static String srcPDF = "srcdata/3.pdf";

    public static void main(String[] args) {

        try {
            file = new FileInputStream(new File("srcExcels/3.xls"));
            workbook = new HSSFWorkbook(file);
            sheet = workbook.getSheetAt(0);
        } catch (Exception ex) {
            ex.printStackTrace();
        }

        try {
            for (int i = 1; i < sheet.getLastRowNum() + 1; i++) {
                pdfCreator(
                        numFinder(i, 0),
                        nameFinder(i, 1),
                        nameFinder(i, 2));

            }

        } catch (Exception ex) {
            ex.getStackTrace();
        }

    }

    private static String numFinder(int i, int cellNum) {
        String name;

        Cell cell = sheet.getRow(i).getCell(cellNum);

        name = Integer.toString((int)cell.getNumericCellValue());

        return name;
    }
    private static String nameFinder(int i, int cellNum) {
        String name;

        Cell cell = sheet.getRow(i).getCell(cellNum);

        name = cell.getStringCellValue();

        return name;
    }

    private static void pdfCreator(String num, String fio, String school) {
        PdfWriter writer = null;
        String filename;

        if (fio.equals(school)){
            school = "";
        }

        if(fio.contains("Шиянова Ирина")){
            System.out.println(fio);
        }

        if (fio.substring(fio.length()-1).contains(" ")){
            fio = fio.substring(0, fio.length()-1);
        }

        if (fio.substring(0, 1).contains(" ")){
            fio = fio.substring(1);
        }

            filename = num + ". Сертификат участника. Вебинар 09.10.2019. " + fio + ".pdf";

        try {
            if(filename.length() > 1){
                if (filename.substring(filename.length() - 1).equals(" ")) {
                    filename = filename.substring(0, filename.lastIndexOf(" "));
                } else if (filename.substring(0, 1).equals(" ")) {
                    filename = filename.substring(1);
                }

                if (filename.contains("\"")) {
                    filename = filename.replaceAll("\"", " ");
                }
                if (filename.contains("№")) {
                    filename = filename.replaceAll("№", " ");
                }

                if (filename.contains("  ")){
                    filename.replaceAll("\\s\\s", " ");
                }
            }

        } catch (Exception ex) {
            ex.printStackTrace();
        }


        try {
            WriterProperties writerProperties = new WriterProperties();
            writerProperties.setStandardEncryption("".getBytes(),
                    "Pusha_i_Pokusenot_Samye_2019_luchshie".getBytes(),
                    EncryptionConstants.ALLOW_PRINTING,
                    EncryptionConstants.STANDARD_ENCRYPTION_128);
            writer = new PdfWriter(dstFolder + filename, writerProperties);
//            writer = new PdfWriter(dstFolder + filename);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

        PdfReader reader = null;
        try {
            reader = new PdfReader(srcPDF);
        } catch (IOException e) {
            e.printStackTrace();
        }

        PdfDocument destpdf = new PdfDocument(writer);

        destpdf.setDefaultPageSize(PageSize.A4.rotate());

        PdfDocument srcPdf = new PdfDocument(reader);

        PdfPage origPage = srcPdf.getPage(1);

        PdfPage page = destpdf.addNewPage();

        PdfFormXObject pageCopy = null;
        try {
            pageCopy = origPage.copyAsFormXObject(destpdf);
        } catch (IOException e) {
            e.printStackTrace();
        }

        PdfCanvas canvas = new PdfCanvas(page);

        canvas.addXObject(pageCopy, 0, 0);

        Document doc = new Document(destpdf);

        PdfFontFactory pdfFontFactory = new PdfFontFactory();


        PdfFont font = null;
        PdfFont fonti = null;
        PdfFont fontb = null;
        try {
            font = pdfFontFactory.createFont("D:\\JavaProjects\\Pusha's tasks\\PDFFormer\\src\\main\\resources\\fonts\\MTCORSVA.TTF", "cp1251", true);
//            fonti = pdfFontFactory.createFont("D:\\JavaProjects\\Pusha's tasks\\PDFFormer\\src\\main\\resources\\fonts\\cambriai.ttf", "cp1251", true);
//            fontb = pdfFontFactory.createFont("D:\\JavaProjects\\Pusha's tasks\\PDFFormer\\src\\main\\resources\\fonts\\cambriab.ttf", "cp1251", true);
//            font = pdfFontFactory.createFont("D:\\JavaProjects\\Pusha's tasks\\PDFFormer\\src\\main\\resources\\fonts\\times.ttf", "cp1251", true);
//            font = pdfFontFactory.createTtcFont("D:\\JavaProjects\\Pusha's tasks\\PDFFormer\\src\\main\\resources\\fonts\\cambria.ttc".getBytes(), 12, "cp1251", true, true);
        } catch (IOException e) {
            e.printStackTrace();
        }

        int fontSize = 15;

        Text line1 = new Text(fio);
        line1.setFontSize(28).setFont(font);
        line1.setRelativePosition(0, 129, 0, 0);

        Text line2 = new Text(school);
        line2.setFont(font);
        line2.setRelativePosition(0, 134, 0, 0);

        if (school.length() > 187) {
            fontSize = 12;
        } else if (school.length() > 140) {
            fontSize = 16;
            line2.setRelativePosition(0, 130, 0, 0);
        } else if (school.length() > 90){
            fontSize = 20;
            line2.setRelativePosition(0, 125, 0, 0);
        } else if (school.length() > 47){
            fontSize = 22;
        }else if (school.length() > 27){
            fontSize = 26;
        }else {
            fontSize = 28;
        }

        line2.setFontSize(fontSize);

        Paragraph par1 = new Paragraph(line1);
        par1.setTextAlignment(TextAlignment.CENTER);

        Paragraph par2 = new Paragraph(line2);
        par2.setTextAlignment(TextAlignment.CENTER);


        doc.add(par1);
        doc.add(par2);

        doc.close();
    }

}
