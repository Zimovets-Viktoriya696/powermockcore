/**
 * Created by Vika on 28.01.2017.
 */

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class Maine {

    public static void main(String... args){
      ArrayList<String>  list = Maine.parse("d:/projects/xlrd/1.xls");
     ArrayList<String> result= Maine.toBinary(Maine.toInt(list));
        System.out.println(result.size());
        for (int i = 0; i < result.size(); i++) {
            System.out.println(result.get(i));
        }
    }


    public static ArrayList<String> parse(String name) {

        String result = "";
        InputStream in = null;
        HSSFWorkbook wb = null;
        ArrayList<String> list = new ArrayList<String>();

        try {
            in = new FileInputStream(name);
            wb = new HSSFWorkbook(in);
        } catch (IOException e) {
            e.printStackTrace();
        }

        Sheet sheet = wb.getSheetAt(0);
        Iterator<Row> it = sheet.iterator();
        for (int i = 6; i< 1018; i++) {
            Row row = sheet.getRow(i);
            Cell cell = row.getCell(1);
                int cellType = cell.getCellType();
                switch (cellType) {
                    case Cell.CELL_TYPE_STRING:
                        result += cell.getStringCellValue();

                             list.add(cell.getRichStringCellValue().toString());
                      //  System.out.println(cell.getStringCellValue() + "mdddddd");


                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        result += cell.getNumericCellValue();
                        list.add(cell.getStringCellValue().toString());


                        break;

                    case Cell.CELL_TYPE_FORMULA:
                        result +=  cell.getNumericCellValue();
                        list.add(cell.getStringCellValue().toString());

                        break;
                    default:

                        break;
                }
            }
            result += "\n";


        return list;
    }

    public static ArrayList<Integer> toInt (ArrayList<String> list){
        ArrayList<Integer> listToInt = new  ArrayList<Integer>();
        for (int i = 0; i < list.size(); i++) {
            try {
                int a = Integer.parseInt(list.get(i));
                listToInt.add(a);
            }
            catch (NumberFormatException e){
                System.out.println("что за херня");
            }
        }
        return listToInt;
    }

    public static ArrayList<String> toBinary(ArrayList<Integer> list){
        ArrayList<String> bin = new ArrayList<String>();
        for (int i = 0; i < list.size(); i++) {
            bin.add(Integer.toBinaryString(list.get(i)));
        }
            return  bin;
    }

}


