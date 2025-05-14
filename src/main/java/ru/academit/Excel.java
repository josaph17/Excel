package ru.academit;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class Excel
{
    public static void main( String[] args ) throws IOException {
        Person person1 = new Person(30, "Rakhmonov", "Shavkat", "+79147652932");
        Person person2 = new Person(38, "Burns", "Gilbert", "+790278543312");
        Person person3 = new Person(36,"Muhammad", "Belal", "+79240001234");

        List<Person> people = new ArrayList<>();
        people.add(person1);
        people.add(person2);
        people.add(person3);

        System.out.println( people );

        Workbook wb = new XSSFWorkbook();

        CreationHelper createHelper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet("new sheet");

        Row row = sheet.createRow(0);
        row.setHeightInPoints(40);

        row.createCell(0).setCellValue(createHelper.createRichTextString("Age"));
        row.createCell(1).setCellValue(createHelper.createRichTextString("Family name"));
        row.createCell(2).setCellValue(createHelper.createRichTextString("Name"));
        row.createCell(3).setCellValue(createHelper.createRichTextString("Number"));

        fillTableData(wb, people, "new sheet");
        setTableStyle(wb);

        for (int i = 0; i < sheet.getRow(0).getPhysicalNumberOfCells(); i++){
            sheet.autoSizeColumn(i);
        }

        try(OutputStream fileOut = new FileOutputStream("C:\\Users\\sayan\\Disk\\!Google Drive for study\\personsTable.xlsx")){
            wb.write(fileOut);
        }

        wb.close();

        System.out.println("Done");
    }

    public static void fillTableData(Workbook wb, List<Person> people, String SheetName) {

        int rowNum = 1;

        for(Person person: people){
            Row row = wb.getSheet(SheetName).createRow(rowNum++);

            row.createCell(0).setCellValue(person.getAge());
            row.createCell(1).setCellValue(person.getFamilyName());
            row.createCell(2).setCellValue(person.getName());
            row.createCell(3).setCellValue(person.getPhoneNumber());
        }

    }

    public static void setTableStyle(Workbook wb) {
        CellStyle commonStyle = wb.createCellStyle();

        commonStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        commonStyle.setAlignment(HorizontalAlignment.CENTER);

        commonStyle.setBorderBottom(BorderStyle.THICK);
        commonStyle.setBottomBorderColor(IndexedColors.GREY_25_PERCENT.getIndex());
        commonStyle.setBorderLeft(BorderStyle.THICK);
        commonStyle.setLeftBorderColor(IndexedColors.GREY_25_PERCENT.getIndex());
        commonStyle.setBorderRight(BorderStyle.THICK);
        commonStyle.setRightBorderColor(IndexedColors.GREY_25_PERCENT.getIndex());
        commonStyle.setBorderTop(BorderStyle.THICK);
        commonStyle.setTopBorderColor(IndexedColors.GREY_25_PERCENT.getIndex());

        CellStyle headerStyle = wb.createCellStyle();
        headerStyle.cloneStyleFrom(commonStyle);
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        CellStyle bodyStyle = wb.createCellStyle();
        bodyStyle.cloneStyleFrom(commonStyle);
        bodyStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        bodyStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        Font headerFont = wb.createFont();
        headerFont.setFontHeightInPoints((short)14);
        headerFont.setFontName("Helvetica");
        headerFont.setItalic(true);
        headerFont.setBold(true);
        headerFont.setColor(IndexedColors.BLACK.getIndex());

        Font bodyFont = wb.createFont();
        bodyFont.setFontHeightInPoints((short)14);
        bodyFont.setFontName("Courier New");
        bodyFont.setItalic(true);

        headerStyle.setFont(headerFont);
        bodyStyle.setFont(bodyFont);

        for (Sheet sheet : wb ) {
            for (Row row : sheet) {

                row.setHeightInPoints(40);
                CellStyle currentStyle = (row.getRowNum() == 0) ? headerStyle : bodyStyle;

                for (Cell cell : row) {
                    cell.setCellStyle(currentStyle);
                }
            }
        }
    }
}
