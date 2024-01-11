package lab1;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Lab1 {

    private final String filePath = System.getProperty("user.dir") + File.separator + "src\\assets\\my_file.xlsx";

    public static void main(String[] args) {
        String[] data = new String[2];
        Lab1 lab1 = new Lab1();
        String format = "%-18s%-18s%-30s%n";
        System.out.printf(format, "a", "b", "X");
        System.out.println("========================================");
        String x;
        for (int i = 1; i < 9; i++) {
            lab1.readData(data, i);

            try {
                if (lab1.isNumber(data[0]) && lab1.isNumber(data[1])) {
                    boolean a = isOutOfRange(data[0]);
                    boolean b = isOutOfRange(data[1]);
                    if (a && b) {
                        x = "'" + data[0] + "' and '" + data[1] + "' are out of range of integer";
                    } else if (!a && b) {
                        x = "'" + data[1] + "' is out of range of integer";
                    } else if (a && !b) {
                        x = "'" + data[0] + "' is out of range of integer";
                    } else {

                        int num1 = Integer.parseInt(data[0]);
                        int num2 = Integer.parseInt(data[1]);
                        x = lab1.add(num1, num2);
                    }
                } else if (!lab1.isNumber(data[0])) {
                    x = "'" + data[0] + "' is not an integer.";
                } else if (!lab1.isNumber(data[1])) {
                    x = "'" + data[1] + "' is not an integer.";
                } else {
                    x = "'" + data[0] + "' and '" + data[1] + "' are not numbers.";
                }
                lab1.writeData(i, 2, x);
                System.out.printf(format, data[0], data[1], x);

            } catch (NumberFormatException e) {
                // Handle exception
            }
        }
    }

    public void readData(String[] result, int row) {

        try {
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(filePath));
            XSSFSheet sh = wb.getSheet("Sheet1");

            for (int i = 0; i < 2; i++) {
                if (sh.getRow(row).getCell(i).getCellType() == CellType.NUMERIC) {
                    result[i] = sh.getRow(row).getCell(i).getRawValue();
                } else {
                    result[i] = sh.getRow(row).getCell(i).getStringCellValue();
                }
            }

        } catch (IOException ex) {
        }
    }

    public void writeData(int row, int col, String data) {
        try {
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(filePath));

            XSSFSheet sh = wb.getSheet("Sheet1");

            sh.getRow(row).createCell(col).setCellValue(data);

            try ( FileOutputStream fileOut = new FileOutputStream(filePath)) {
                wb.write(fileOut);
            }
            wb.close();

        } catch (IOException ex) {
//            Logger.getLogger(Lab1.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    public String add(int a, int b) {
        long result = 0;
        int x = 0;
        for (int i = 0; i < 10000; i++) {
            for (int j = 0; j < 10000; j++) {
                result += a + b;
                x += a + b;
                if (isOutOfRange(result)) {
                    return x + " Last X is out of range of Integer";
                }
            }
        }
        return x + "";
    }

    public boolean isNumber(String input) {
        return input.matches("-?\\d+?");
    }

    public static boolean isOutOfRange(String input) {
        try {
            int parsedValue = Integer.parseInt(input);
            return parsedValue < Integer.MIN_VALUE || parsedValue > Integer.MAX_VALUE;
        } catch (NumberFormatException e) {
            return true;
        }
    }

    public static boolean isOutOfRange(long input) {
        return input < Integer.MIN_VALUE || input > Integer.MAX_VALUE;
    }
}
