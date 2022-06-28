package ru.read.file;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

public class ExMail {

    private static String [] caps = new String[]{"Дата", "Контрагент",
            "Адрес доставки \n" + "(обязательно телефон отправителя и получателя!)", "Номер документа",
            "Исполнитель", "Примечание", "Трек-номер"};

    private static ArrayList<Object> list = new ArrayList<>();
    private static Workbook book_res = new XSSFWorkbook();
    private static Sheet sheet_res = book_res.createSheet("Газпром ExMail");

    private static String fileRes = "C:\\Users\\Илья\\Desktop\\Илья\\Газпром ExMail\\! 1 Реестр Газпром ExMail.xlsx";
    private static File folder = new File("C:\\Users\\Илья\\Desktop\\Илья\\Газпром ExMail\\");

    private static File[] listOfFiles = folder.listFiles();

    private static Integer len_row = 1;

    public static void main(String[] args) throws IOException {

        createCaps(caps);

        for (File file : listOfFiles) {
            if (file.isFile() && file.getName().matches(".*\\.xlsx")) {
                String ways = folder + "\\" + file.getName();
                int cnt = 0;
                parse(ways, cnt);
                createTable();
            }
        }

        FileOutputStream out = new FileOutputStream(fileRes);
        book_res.write(out);
        out.close();
        book_res.close();
        dataTransfer();
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

    public static void parse(String fileName, int cnt) throws IOException {
        //инициализируем потоки
        InputStream inputStream = null;
        XSSFWorkbook workBook = null;
        try {
            inputStream = new FileInputStream(fileName);
            workBook = new XSSFWorkbook(inputStream);
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
            cnt = 0;
            while (cells.hasNext()) {
                Cell cell = cells.next();
                for (int i = 0; i < caps.length; i++){
                    if (cell.getColumnIndex() == i) {
                        cnt++;
                        if (cell.getCellType() == CellType.STRING) {
                            list.add(cell.getStringCellValue());
                        } else if (cell.getCellType() == CellType.NUMERIC) {
                            if (DateUtil.isCellDateFormatted(cell)){
                                SimpleDateFormat df = new SimpleDateFormat("dd.MM.yyyy");
                                Date date = cell.getDateCellValue();
                                String cellValue = df.format(date);
                                list.add(cellValue);
                            }
                            else {
                                list.add(cell.getNumericCellValue());
                            }
                        }
                    }
                }
            }
            for (int i = 0; i < caps.length - cnt; i++){
                list.add("Пусто");
            }
        }
        inputStream.close();
        workBook.close();
    }

    private static void dataTransfer() throws IOException {
        SimpleDateFormat formatter = new SimpleDateFormat("dd.MM.yyyy");
        Date date = new Date();
        String name_path_xlsx = formatter.format(date) + " Газпром ExMail";
        String name_path_other = formatter.format(date) + " Газпром ExMail (не .xlsx файлы)";

        File source = new File("C:\\Users\\Илья\\Desktop\\Илья\\Газпром ExMail\\");
        File dest = new File("C:\\Users\\Илья\\Desktop\\Илья\\" + name_path_xlsx);
        File other = new File("C:\\Users\\Илья\\Desktop\\Илья\\" + name_path_other);


        try {
            FileUtils.copyDirectory(source, dest);
        } catch (IOException e) {
            e.printStackTrace();
        }

        FileUtils.cleanDirectory(source);

        File registry = new File("C:\\Users\\Илья\\Desktop\\Илья\\" + name_path_xlsx + "\\! 1 Реестр Газпром ExMail.xlsx");
        FileUtils.copyFileToDirectory(registry, source);
        registry.delete();

        File[] destOfFiles = dest.listFiles();
        for (File file : destOfFiles) {
            if (!(file.getName().matches(".*\\.xlsx"))) {
                FileUtils.copyFileToDirectory(file, other);
                file.delete();
            }
        }
    }

    private static void createTable() {

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
        while (length > caps.length - 1){
            Row row = sheet_res.createRow(len_row);
            for (int k = index; k < list.size() - length + caps.length; k++){
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

            length -= caps.length;
            index += caps.length;
            len_row = len_row + 1;
        }
        list.clear();
    }
}