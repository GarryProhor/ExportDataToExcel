package exportdatatoexel.utils;

import exportdatatoexel.entity.Customer;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;
import org.springframework.util.ObjectUtils;
import org.springframework.util.ResourceUtils;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.lang.reflect.Field;
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
        insertFieldNameAsTitleToWorkbook(ExportConfig.customerExport.getCellExportConfigList(), newSheet, tittleCellStyle);
        //insert data of fieldName to excel

        //return
    }

    private static <T> void insertDataToWorkbook(Workbook workbook,
                                                 ExportConfig exportConfig,
                                                 List<T> datas,
                                                 XSSFCellStyle dataCellStyle){

        int startRowIndex = exportConfig.getStartRow();//2

        int sheetIndex = exportConfig.getSheetIndex();//1

        Class clazz = exportConfig.getDataClazz();

        List<CellConfig> cellConfigs = exportConfig.getCellExportConfigList();

        Sheet sheet = workbook.getSheetAt(sheetIndex);

        int currentRowIndex = startRowIndex;

        for(T data : datas){
            Row currentRow = sheet.getRow(currentRowIndex);
            if(ObjectUtils.isEmpty(currentRow)){
                currentRow = sheet.createRow(currentRowIndex);
            }

            // insert data to row
        }
    }
    private static <T> void insertFieldNameAsTitleToWorkbook(List<CellConfig> cellConfigs,
                                                             Sheet sheet,
                                                             XSSFCellStyle titleCellStyle){
        //title -> first row of excel -> get top row
        int currentRow = sheet.getTopRow();

        //create row
        Row row = sheet.createRow(currentRow);
        int i = 0;

        //resize fix text in each cell
        sheet.autoSizeColumn(currentRow);

        //insert field name to cell
        for(CellConfig cellConfig : cellConfigs){
            Cell currentCell = row.createCell(i);
            String fieldName = cellConfig.getFieldName();
            currentCell.setCellValue(fieldName);
            currentCell.setCellStyle(titleCellStyle);
            sheet.autoSizeColumn(i);
            i++;
        }
    }

    private static <T> void insertDataToCell(T data, Row currentRow, List<CellConfig> cellConfigs,
                                             Class clazz, Sheet sheet, XSSFCellStyle dataStyle){
        for(CellConfig cellConfig : cellConfigs){
            Cell currentCell = currentRow.getCell(cellConfig.getColumnIndex());
            if(ObjectUtils.isEmpty(currentCell)){
                currentCell = currentRow.createCell(cellConfig.getColumnIndex());
            }

            //get data for cell
            String cellValue = getCellValue(data, cellConfig, clazz);
        }
    }
    private static <T> String getCellValue(T data, CellConfig cellConfig, Class clazz){
        String fieldName = cellConfig.getFieldName();

        try {
            Field field = getDeclaredField(clazz, fieldName);
        }catch (Exception e){

        }
    }
    private static Field getDeclaredField(Class clazz, String fieldName){
        if(ObjectUtils.isEmpty(clazz) || ObjectUtils.isEmpty(fieldName)){
            return  null;
        }
        do{
            try {
                Field field = clazz.getDeclaredField(fieldName);
                field.setAccessible(true);
            } catch (NoSuchFieldException e) {
                ;log.info("" + e);
            }
        }while ((clazz = clazz.getSuperclass()) !=null);
        return null;
    }

}
