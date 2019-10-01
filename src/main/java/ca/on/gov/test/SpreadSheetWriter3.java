package ca.on.gov.test;

import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.Iterator;
import java.util.Random;

public class SpreadSheetWriter3 {
    private static File completedDirectory;
    private static File spreadsheetsDirectory;

    public static void main(String[] args)
    {
        boolean sequential = true;

        //sequential = Boolean.parseBoolean(args[0]);

        int seqInt = 100;

        completedDirectory = new File("Completed Spreadsheets");
        spreadsheetsDirectory = new File("Spreadsheets");
        if (!verifyDirectories()) {
            System.exit(0);
        }
        File[] excelFiles = spreadsheetsDirectory.listFiles();
        for (File file : excelFiles)
        {
            int numerictype = 0;
            int stringtype = 0;
            int blanktype = 0;

            System.out.println("Working on: " + file.getName());

            String[] splitFileAbsPath = file.getAbsolutePath().split("\\.");
            if ((splitFileAbsPath[(splitFileAbsPath.length - 1)].equals("xls")) ||
                    (splitFileAbsPath[(splitFileAbsPath.length - 1)].equals("xlsx")) ||
                    (splitFileAbsPath[(splitFileAbsPath.length - 1)].equals("xlsm"))) {
                try
                {
                    Workbook workbook = WorkbookFactory.create(new FileInputStream(file));

                    Sheet[] sheets = new Sheet[workbook.getNumberOfSheets()];
                    FileOutputStream fileOut = new FileOutputStream("Completed Spreadsheets/" + file.getName());
                    for (int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++) {
                        sheets[sheetNum] = workbook.getSheetAt(sheetNum);
                    }
                    Iterator localIterator1;
                    Row row;
                    for (Sheet sheet : sheets) {
                      //  if(!sheet.getProtect())
                        {
                            for (localIterator1 = sheet.iterator(); localIterator1.hasNext(); ) {
                                row = (Row) localIterator1.next();
                                for (Cell cell : row) {
                                    try {
                                        if (((!cell.getCellStyle().getLocked()) || cell.getCellType() != CellType.ERROR|| cell.getCellType() !=CellType.FORMULA)) {
                                            if (cell.getCellType() == CellType.NUMERIC) {
                                                if (cell.getNumericCellValue() == 0.0D) {
                                                    if (sequential) {
                                                        cell.setCellValue(seqInt);
                                                    } else {
                                                        cell.setCellValue(randInt(1, 10000));
                                                    }
                                                    numerictype++;
                                                    seqInt++;
                                                }
                                            } else if (cell.getCellType() == CellType.STRING) {
                                                if (cell.getStringCellValue().equalsIgnoreCase("item description") ||
                                                        cell.getStringCellValue().equalsIgnoreCase("enter comments here") ||
                                                        cell.getStringCellValue().equalsIgnoreCase("$0")
                                                ) {
                                                    if (sequential) {
                                                        cell.setCellValue(seqInt);
                                                    } else {
                                                        cell.setCellValue(randInt(1, 10000));
                                                    }
                                                    stringtype++;
                                                    seqInt++;
                                                }
                                            } else if (cell.getCellType() == CellType.BLANK) {
                                                if (sequential) {
                                                    cell.setCellValue(seqInt);
                                                } else {
                                                    cell.setCellValue(randInt(1, 10000));
                                                }
                                                blanktype++;
                                                seqInt++;
                                            }
                                        } else {
                                            if (cell.getCellType() == CellType.ERROR)
                                                System.out.println("Error Cell");
                                        }
                                    } catch (Exception localException) {
                                        localException.printStackTrace();
                                    }
                                }
                            }
                        }
                    }
                    try
                    {
                        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                        evaluator.evaluateAll();
                    }
                    catch (Exception localException2) {}
                    workbook.write(fileOut);
                    fileOut.flush();
                    fileOut.close();

                    System.out.println("Completed: " + file.getName());

                    System.out.println("Numeric: " + numerictype);
                    System.out.println("String: " + stringtype);
                    System.out.println("Blank: " + blanktype);
                }
                catch (FileNotFoundException e)
                {
                    System.out.println(e.getMessage());
                }
                catch (IOException e)
                {
                    System.out.println(e.getMessage());
                }
            }
        }
    }

    public static int randInt(int min, int max)
    {
        Random rand = new Random();

        int randomNum = rand.nextInt(max - min + 1) + min;

        return randomNum;
    }

    private static boolean verifyDirectories() {
        boolean result = true;
        if (!completedDirectory.exists()) {
            try {
                completedDirectory.mkdir();
            } catch (SecurityException e) {
                System.out.println(e.getMessage());
                result = false;
            }
        }
        if (!spreadsheetsDirectory.exists()) {
            try {
                spreadsheetsDirectory.mkdir();
            } catch (SecurityException e) {
                System.out.println(e.getMessage());
                result = false;
            }
        }
        return result;
    }
}
