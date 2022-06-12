package ru.read.file;

import org.apache.commons.collections4.MultiValuedMap;
import org.apache.commons.collections4.multimap.ArrayListValuedHashMap;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.util.Iterator;

public class Main {

    private static MultiValuedMap<String, Double> map = new ArrayListValuedHashMap<>();
    private static String ways = "/Users/leonid/Desktop/Мама_2/Физкультура-ТЗ №%d-оценки.xls";
    private static String fileRes = "/Users/leonid/Desktop/Мама_2/Результат 3 курс.xls";
    private static final int start = 9;
    private static final int count = 25;

    public static void main(String[] args) throws IOException {
        for (int i = start; i <= count; i++) {
            parse(String.format(ways, i), i - start);
        }
        addZeroAll();
//        for (String key : map.keySet()) {
//            System.out.print(key + " -> ");
//            map.get(key).forEach(x -> System.out.print(x + " | "));
//            System.out.println();
//        }
        createTable(fileRes);
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
            StringBuilder studentInfo = new StringBuilder();
            Iterator<Cell> cells = row.iterator();
            while (cells.hasNext()) {
                Cell cell = cells.next();
                if (cell.getColumnIndex() < 3) {
                    studentInfo.append(cell.getStringCellValue().trim()).append(" ");
                } else if (cell.getColumnIndex() == 3) {
                    try {
                        String key = studentInfo.toString().trim();
                        Double value = Double.valueOf(cell.getStringCellValue().replace(",", "."));
                        addZero(key, iterate);
                        map.get(key).add(value);
                    } catch (IllegalStateException e) {
                        System.out.println(cell.getStringCellValue() + " -> Ошибка");
                    }
                }
            }
        }
    }

    private static void addZero(String key, int iterate) {
        while (map.get(key).size() < iterate) {
            map.get(key).add(0.d);
        }
    }

    private static void addZeroAll() {
        for (String key : map.keySet()) {
            while (map.get(key).size() < count - start) {
                map.get(key).add(0.d);
            }
        }
    }

    private static void createTable(String file) throws IOException {
        Workbook book = new HSSFWorkbook();
        Sheet sheet = book.createSheet("Birthdays");
        int indexRow = 0;
        int indexCol = 3;

        for (String key : map.keySet()) {
            Row row = sheet.createRow(indexRow++);
            String[] studentInfo = key.split(" ");
            for (int i = 0; i < studentInfo.length; i++) {
                Cell name = row.createCell(i);
                name.setCellValue(studentInfo[i]);
            }

            for (Double value : map.get(key)) {
                Cell valueCol = row.createCell(indexCol++);
                valueCol.setCellValue(value);
            }
//            Cell sum = row.createCell(indexCol);
//            sum.setCellValue(map.get(key).stream().mapToDouble(x -> x).sum() / count);
            // Меняем размер столбца
            indexCol = 3;
        }
        // Записываем всё в файл
        book.write(new FileOutputStream(file));
        book.close();
    }
}