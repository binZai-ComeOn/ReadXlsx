package com.example.readxlsx.utils;

import android.content.ContentValues;
import android.content.Context;
import android.database.Cursor;
import android.database.sqlite.SQLiteDatabase;
import android.net.Uri;
import android.widget.Toast;

import com.example.readxlsx.sqlite.MySQLHelp;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelUtil {
    List<Map<Integer, Object>> dataList = new ArrayList<>();
    Workbook workbook;
    String excelStr;
    InputStream inputStream;

    MySQLHelp mySQLHelp;
    ContentValues contentValues;
    Cursor cursor;

    public List<Map<Integer, Object>> readExcel(Context context, Uri fileUri, String strFileUri) {
        mySQLHelp = new MySQLHelp(context, "mydb.db", null, 1);
        SQLiteDatabase writableDatabase = mySQLHelp.getWritableDatabase();
        excelStr = strFileUri.substring(strFileUri.lastIndexOf("."));
        try {
            inputStream = context.getContentResolver().openInputStream(fileUri);
            if (excelStr.equals(".xlsx")) workbook = new XSSFWorkbook(inputStream);
            else if (excelStr.equals(".xls")) workbook = new HSSFWorkbook(inputStream);
            else workbook = null;
            if (workbook != null) {
                Sheet sheetAt = workbook.getSheetAt(0);
                Row row = sheetAt.getRow(0);
                int physicalNumberOfCells = row.getPhysicalNumberOfCells();//获取实际单元格数
                Map<Integer, Object> map = new HashMap<>();
                for (int i = 0; i < physicalNumberOfCells; i++) {//将标题存储到map
                    Object cellFormatValue = getCellFormatValue(row.getCell(i));
                    map.put(i, cellFormatValue);
                }
                dataList.add(map);
                int physicalNumberOfRows = sheetAt.getPhysicalNumberOfRows();//获取最大行数
                int size = map.size();//获取最大列数
                contentValues = new ContentValues();
                for (int i = 1; i < physicalNumberOfRows; i++) {
                    Map<Integer, Object> map1 = new HashMap<>();
                    Row row1 = sheetAt.getRow(i);
                    if (!row1.equals(null)) {
                        for (int j = 0; j < size; j++) {
                            Object cellFormatValue = getCellFormatValue(row1.getCell(j));
                            map1.put(j, cellFormatValue);
                            System.out.println(j);
                        }
                        contentValues.put("materialID", (String) map1.get(0));
                        contentValues.put("materialEncoding", (String) map1.get(1));
                        contentValues.put("materialName", (String) map1.get(2));
                        contentValues.put("materialModel", (String) map1.get(3));
                        contentValues.put("materialSize", (String) map1.get(4));
                        contentValues.put("unit", (String) map1.get(5));
                        contentValues.put("price", (String) map1.get(6));
                        contentValues.put("count", (String) map1.get(7));
                        contentValues.put("manufacturers", (String) map1.get(8));
                        contentValues.put("type", (String) map1.get(9));
                        contentValues.put("receiptor", (String) map1.get(10));
                        contentValues.put("storagelocation", (String) map1.get(11));
                        contentValues.put("materialState", (String) map1.get(12));
                        writableDatabase.insert("module", null, contentValues);
                    } else break;
                    dataList.add(map1);
                }
                writableDatabase.close();
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return dataList;
    }

    /*
     * 获取单元格格式值
     * */
    private static Object getCellFormatValue(Cell cell) {
        Object cellValue;
        if (cell != null) {
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_BOOLEAN:
                    cellValue = cell.getBooleanCellValue();
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    cellValue = String.valueOf(cell.getNumericCellValue());
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        cellValue = cell.getDateCellValue();
                    } else {
                        cellValue = cell.getNumericCellValue();
                    }
                    break;
                case Cell.CELL_TYPE_STRING:
                    cellValue = cell.getStringCellValue();
                    break;
                default:
                    cellValue = "";
            }
        } else {
            cellValue = "";
        }
        return cellValue;
    }

    public void getDataAndSave(Context context,Uri uri) {
        ArrayList<Map<Integer,Object>> arrayList = new ArrayList<>();
        Map<Integer,Object> m = new HashMap<>();
        m.put(0,"物料ID");
        m.put(1,"物料编码");
        m.put(2,"名称");
        m.put(3,"编号");
        m.put(4,"规格");
        m.put(5,"单位");
        m.put(6,"单价");
        m.put(7,"数量");
        m.put(8,"厂家");
        m.put(9,"类别");
        m.put(10,"经手人");
        m.put(11,"存放地点");
        m.put(12,"状态");
        arrayList.add(m);

        mySQLHelp = new MySQLHelp(context, "mydb.db", null, 1);
        SQLiteDatabase readableDatabase = mySQLHelp.getReadableDatabase();
        cursor = readableDatabase.rawQuery("select * from module", null);
        while (cursor.moveToNext()) {
            Map<Integer,Object> map = new HashMap<>();
            String materialID = cursor.getString(cursor.getColumnIndex("materialID"));
            String materialEncoding = cursor.getString(cursor.getColumnIndex("materialEncoding"));
            String materialName = cursor.getString(cursor.getColumnIndex("materialName"));
            String materialModel = cursor.getString(cursor.getColumnIndex("materialModel"));
            String materialSize = cursor.getString(cursor.getColumnIndex("materialSize"));
            String unit = cursor.getString(cursor.getColumnIndex("unit"));
            String price = cursor.getString(cursor.getColumnIndex("price"));
            String count = cursor.getString(cursor.getColumnIndex("count"));
            String manufacturers = cursor.getString(cursor.getColumnIndex("manufacturers"));
            String type = cursor.getString(cursor.getColumnIndex("type"));
            String receiptor = cursor.getString(cursor.getColumnIndex("receiptor"));
            String storagelocation = cursor.getString(cursor.getColumnIndex("storagelocation"));
            String materialState = cursor.getString(cursor.getColumnIndex("materialState"));

            map.put(0,materialID);
            map.put(1,materialEncoding);
            map.put(2,materialName);
            map.put(3,materialModel);
            map.put(4,materialSize);
            map.put(5,unit);
            map.put(6,price);
            map.put(7,count);
            map.put(8,manufacturers);
            map.put(9,type);
            map.put(10,receiptor);
            map.put(11,storagelocation);
            map.put(12,materialState);
            arrayList.add(map);
        }
        try {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet(WorkbookUtil.createSafeSheetName("Sheet1"));
            Cell cell;
            int size = arrayList.get(0).size();
            for (int i = 0;i < arrayList.size();i++){
                Row row = sheet.createRow(i);
                Map<Integer, Object> map = arrayList.get(i);
                for (int j = 0;j < size;j++){
                    cell = row.createCell(j);
                    cell.setCellValue((String) map.get(j));
                    System.out.println(map.get(j)+"\n");
                }
            }
            OutputStream outputStream = context.getContentResolver().openOutputStream(uri);
            workbook.write(outputStream);
            outputStream.flush();
            outputStream.close();
            Toast.makeText(context, "另存成功", Toast.LENGTH_SHORT).show();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}