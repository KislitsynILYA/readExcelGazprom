package ru.read.file;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;

public class RussianPost {

    private static String [] caps = new String[]{"Дата", "Контрагент",
            "Адрес доставки \n" + "(обязательно индекс!)",
            "Номер документа", "Исполнитель", "Примечание", "Вес", "Трек-номер"};

    private static ArrayList<Object> list = new ArrayList<>();
    private static Workbook book_res = new HSSFWorkbook();
    private static Sheet sheet_res = book_res.createSheet("Газпром Почта России");

    private static String fileRes = "C:\\Users\\Илья\\Desktop\\Илья\\Газпром Почта России\\1 Реестр Газпром Почта России.xls";
    private static File folder = new File("C:\\Users\\Илья\\Desktop\\Илья\\Газпром Почта России\\");

    private static File[] listOfFiles = folder.listFiles();

    private static Integer len_row = 1;

    public static void main(String[] args) throws IOException {

        createCaps(caps);

        for (File file : listOfFiles) {
            if (file.isFile()) {
                String ways = folder + "\\" + file.getName();
                parse(ways);
                System.out.println(list.size());
                createTable(fileRes);
            }
        }
        //dataTranser(folder);
    }

    public static void createCaps (String [] caps){

        CellStyle style = book_res.createCellStyle();
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setBorderBottom(BorderStyle.MEDIUM);
        style.setBorderLeft(BorderStyle.MEDIUM);
        style.setBorderRight(BorderStyle.MEDIUM);
        style.setBorderTop(BorderStyle.MEDIUM);

        Font font = book_res.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 12);
        style.setFont(font);

        Row row = sheet_res.createRow(0);
        for (int i = 0; i < caps.length; i++ ){
            Cell name = row.createCell(i);
            name.setCellValue(caps[i]);
            name.setCellStyle(style);
        }
    }

    public static void parse(String fileName) {
        //инициализируем потоки
        InputStream inputStream = null;
        HSSFWorkbook workBook = null;
        try {
            inputStream = new FileInputStream(fileName);
            workBook = new HSSFWorkbook(inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        //разбираем первый лист входного файла на объектную модель
        Sheet sheet = workBook.getSheetAt(0);
        Iterator<Row> it = sheet.iterator();
        it.next();
        //проходим по всему листу
        while (it.hasNext()) {
            Row row = it.next();
            Iterator<Cell> cells = row.iterator();
            while (cells.hasNext()) {
                Cell cell = cells.next();
                for (int i = 0; i < caps.length; i++){
                    if (cell.getColumnIndex() == i) {
                        if (cell.getCellType() == CellType.STRING) {
                            list.add(cell.getStringCellValue());
                        } else if (cell.getCellType() == CellType.NUMERIC) {
                            list.add(cell.getNumericCellValue());
                        } else if (cell.getCellType() == CellType.BLANK){
                            list.add("Пусто");
                        }
                    }
                }
            }
        }
    }

    private static void dataTranser(File folder) {
        SimpleDateFormat formatter = new SimpleDateFormat("dd.MM.yyyy");
        Date date = new Date();
        String name_path = formatter.format(date) + " Газпром Почта России";

        File source = new File("C:\\Users\\Илья\\Desktop\\Илья\\Газпром Почта России\\");
        File dest = new File("C:\\Users\\Илья\\Desktop\\Илья\\" + name_path);
        try {
            FileUtils.copyDirectory(source, dest);
        } catch (IOException e) {
            e.printStackTrace();
        }
        String [] source_arr = source.list();
        for(String s: source_arr){
            File currentFile = new File(source.getPath(),s);
            if (currentFile.equals(new File(source.getPath(),
                    "C:\\Users\\Илья\\Desktop\\Илья\\Газпром Почта России\\1 Реестр Газпром Почта России.xls"))){
                continue;
            }
            currentFile.delete();
        }

        File del = new File("C:\\Users\\Илья\\Desktop\\Илья\\" + name_path + "\\1 Реестр Газпром Почта России.xls");
        del.delete();
    }

    private static void createTable(String file) throws IOException {

        for (int j = 0; j < list.size(); j++){
            if (list.get(j).equals("")){
                list.remove(j);
            }
        }

        CellStyle style1 = book_res.createCellStyle();
        style1.setBorderBottom(BorderStyle.THIN);
        style1.setBorderLeft(BorderStyle.THIN);
        style1.setBorderRight(BorderStyle.THIN);
        style1.setBorderTop(BorderStyle.THIN);

        CellStyle style2 = book_res.createCellStyle();
        style2.setBorderBottom(BorderStyle.THIN);
        style2.setBorderLeft(BorderStyle.THIN);
        style2.setBorderRight(BorderStyle.THIN);
        style2.setBorderTop(BorderStyle.THIN);
        Font font2 = book_res.createFont();
        font2.setBold(true);
        font2.setColor(IndexedColors.RED.getIndex());
        style2.setFont(font2);

        int index = 0;
        int length = list.size();
        while (length > 7){
            Row row = sheet_res.createRow(len_row);
            for (int k = index; k < list.size() - length + 8; k++){
                Cell name = row.createCell(k - index);
                if (list.get(k) instanceof String){
                    name.setCellValue((String) list.get(k));
                }
                else if (list.get(k) instanceof Double){
                    name.setCellValue((Double) list.get(k));
                }

                if (list.get(k).equals("Пусто")){
                    name.setCellStyle(style2);
                }
                else {
                    name.setCellStyle(style1);
                }
                for (int l = 0; l < caps.length; l++) sheet_res.autoSizeColumn(l);
            }

            length -= 8;
            index += 8;
            len_row = len_row + 1;
        }

        list.clear();
        book_res.write(new FileOutputStream(file));
        book_res.close();
    }
}
