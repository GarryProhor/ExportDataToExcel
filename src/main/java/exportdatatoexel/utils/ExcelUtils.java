package exportdatatoexel.utils;

import exportdatatoexel.entity.Customer;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;
import org.springframework.util.ResourceUtils;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.util.List;

import static exportdatatoexel.utils.FileFactory.PATH_TEMPLATE;

@Slf4j
@Component
public class ExcelUtils {

    public static ByteArrayInputStream exportCustomer(List<Customer> customers, String fileName) throws Exception {
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        //get file -> not found -> create file
        File file;
        FileInputStream fileInputStream;

        try{
            file = ResourceUtils.getFile(PATH_TEMPLATE + fileName);
            fileInputStream = new FileInputStream(file);
        }catch (Exception e){
            log.info("FILE NOT FOUND");
            file = FileFactory.createFile(fileName,xssfWorkbook);
            fileInputStream = new FileInputStream(file);
        }

        //create freeze pane in excel file
        XSSFSheet newSheet = xssfWorkbook.createSheet("sheet1");
        newSheet.createFreezePane(4, 2, 4, 2);

        //create font for title
        XSSFFont titleFont = xssfWorkbook.createFont();
        titleFont.setFontName("Arial");
        titleFont.setBold(true);
        titleFont.setFontHeightInPoints((short) 14);

        //create style for cell of title and apply font to cell
        XSSFCellStyle tittleCellStyle = xssfWorkbook.createCellStyle();
        tittleCellStyle.setAlignment(HorizontalAlignment.CENTER);
        tittleCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        tittleCellStyle.setFillForegroundColor(IndexedColors.SKY_BLUE.index);
        tittleCellStyle.setBorderBottom(BorderStyle.MEDIUM);
        tittleCellStyle.setBorderLeft(BorderStyle.MEDIUM);
        tittleCellStyle.setBorderRight(BorderStyle.MEDIUM);
        tittleCellStyle.setBorderTop(BorderStyle.MEDIUM);
        tittleCellStyle.setFont(titleFont);
        tittleCellStyle.setWrapText(true);

        //create font for data
        XSSFFont dataFont = xssfWorkbook.createFont();
        dataFont.setFontName("Arial");
        dataFont.setBold(false);
        dataFont.setFontHeightInPoints((short) 10);

        //create style for cell data and apply font data to cell
        XSSFCellStyle dataCellStyle = xssfWorkbook.createCellStyle();
        dataCellStyle.setAlignment(HorizontalAlignment.CENTER);
        dataCellStyle.setBorderBottom(BorderStyle.THIN);
        dataCellStyle.setBorderLeft(BorderStyle.THIN);
        dataCellStyle.setBorderRight(BorderStyle.THIN);
        dataCellStyle.setBorderTop(BorderStyle.THIN);
        dataCellStyle.setFont(dataFont);
        dataCellStyle.setWrapText(true);

        //insert fieldName as title to excel

        //insert data of fieldName to excel

        //return
    }
}
