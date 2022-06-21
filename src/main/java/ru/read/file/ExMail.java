package ru.read.file;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;

public class ExMail {

    private static String [] caps = new String[]{"Дата", "Контрагент", "Адрес доставки", "Номер документа",
            "Исполнитель", "Накладная"};

    private static ArrayList<String> list = new ArrayList<>();
    private static Workbook book_res = new HSSFWorkbook();
    private static Sheet sheet_res = book_res.createSheet("Газпром");

    private static String fileRes = "C:\\Users\\Илья\\Desktop\\Илья\\Газпром ExMail\\1 Реестр Газпром ExMail.xls";
    private static File folder = new File("C:\\Users\\Илья\\Desktop\\Илья\\Газпром ExMail\\");

    private static File[] listOfFiles = folder.listFiles();

    private static Integer len_row = 1;

    public static void main(String[] args) throws IOException {

        Row row = sheet_res.createRow(0);
        for (int i = 0; i < caps.length; i++ ){
            Cell name = row.createCell(i);
            name.setCellValue(caps[i]);
        }


        for (File file : listOfFiles) {
            if (file.isFile()) {
                String ways = folder + "\\" + file.getName();
                parse(ways, 0);
                System.out.println(list.size());
                createTable(fileRes);
            }
        }
        dataTranser(folder);
    }

    public static void parse(String fileName, int iterate) {
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
                if (cell.getColumnIndex() == 0) {
                    list.add(cell.getStringCellValue());
                }
                else if (cell.getColumnIndex() == 1) {
                    list.add(cell.getStringCellValue());;
                }
                else if (cell.getColumnIndex() == 2) {
                    list.add(cell.getStringCellValue());
                }
                else if (cell.getColumnIndex() == 3) {
                    list.add(cell.getStringCellValue());
                }
                else if (cell.getColumnIndex() == 4) {
                    list.add(cell.getStringCellValue());
                }
            }
        }
    }

    private static void dataTranser(File folder) {
        SimpleDateFormat formatter = new SimpleDateFormat("dd.MM.yyyy");
        Date date = new Date();
        String name_path = formatter.format(date) + " Газпром ExMail";

        File source = new File("C:\\Users\\Илья\\Desktop\\Илья\\Газпром ExMail\\");
        File dest = new File("C:\\Users\\Илья\\Desktop\\Илья\\" + name_path);
        try {
            FileUtils.copyDirectory(source, dest);
        } catch (IOException e) {
            e.printStackTrace();
        }
        String [] entries = source.list();
        for(String s: entries){
            File currentFile = new File(source.getPath(),s);
            if (currentFile.equals(new File(source.getPath(),
                    "C:\\Users\\Илья\\Desktop\\Илья\\Газпром ExMail\\1 Реестр Газпром Почта России.xls"))){
                continue;
            }
            currentFile.delete();
        }

        File del = new File("C:\\Users\\Илья\\Desktop\\Илья\\" + name_path + "\\1 Реестр Газпром ExMail.xls");
        del.delete();

    }

    private static void createTable(String file) throws IOException {

        for (int j = 0; j < list.size(); j++){
            if (list.get(j).equals("")){
                list.remove(j);
            }
        }

        int index = 0;
        int length = list.size();
        while (length > 4){
            Row row = sheet_res.createRow(len_row);
            for (int k = index; k < list.size() - length + 5; k++){
                Cell name = row.createCell(k - index);
                name.setCellValue(list.get(k));
            }

            length -= 5;
            index += 5;
            len_row = len_row + 1;
        }

        list.clear();
        book_res.write(new FileOutputStream(file));
        book_res.close();
    }
}
