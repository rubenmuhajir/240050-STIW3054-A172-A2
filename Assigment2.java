package com.mycompany.assignment2;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.Writer;
import java.util.Iterator;
import java.util.Scanner;
import java.util.StringTokenizer;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/**
 *
 * @author USER
 */

public class Assigment2 {

    public static void run() {

        Writer w = null;
        boolean LineOut = true;

        try {

            DataFormatter dataformat = new DataFormatter();
            FileInputStream excel = new FileInputStream(new File("C:\\Users\\USER\\Desktop\\studList.xlsx"));
            Workbook workbook = new XSSFWorkbook(excel);
            Sheet data = workbook.getSheetAt(0);
            Iterator<Row> iterator = data.iterator();

            File file = new File("C:\\Users\\USER\\Desktop\\240050.md");
            w = new BufferedWriter(new FileWriter(file));

            while (iterator.hasNext()) {

                Row row = iterator.next();
                Iterator<Cell> cellIterator = row.iterator();

                while (cellIterator.hasNext()) {

                    Cell cell = cellIterator.next();
                    String value = dataformat.formatCellValue(cell);

                    System.out.print(value + "|");

                    w.write(value + "|");
                }
                System.out.println();
                w.write("\n");
                if (LineOut == true) {
                    w.write("---|---|---|---|\n");
                    LineOut = false;
                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        try {
            if (w != null) {
                w.close();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


public static void main(String[] args) throws IOException{
        File file = new File("C:\\Users\\USER\\Desktop\\240050.md");
        Scanner scan = new Scanner(file);
            
        String lines = null;
        int numL =0;
        int w = 0;
        int c = 0;

        while(scan.hasNextLine())  {
            numL++;
            lines = scan.nextLine();
            c+= lines.length();
            w+= new StringTokenizer(lines, " ,").countTokens();
    }

            System.out.println("The total of lines: " + numL);
            System.out.println("The total of words: " + w);
            System.out.println("The total of characters: " + c);
}
}