package ca.on.gov.test.ssw.poi;

import ca.on.gov.test.ssw.SpreadSheetWriter3;
import ca.on.gov.test.ssw.util.Helper;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.Iterator;
import java.util.List;

public class ExcelWorker implements Runnable {

    private int seqInt = 100;

    private File spreadsheet;
    private File completedDirectory;

    private boolean sequential;

    public ExcelWorker(File spreadsheet, boolean sequential) {
        this.spreadsheet = spreadsheet;
        this.sequential = sequential;

        this.completedDirectory = new File(SpreadSheetWriter3.Directory.COMPLETED.toString());
    }

    @Override
    public void run() {
        int numerictype = 0;
        int stringtype = 0;
        int blanktype = 0;

        System.out.println("Working on: " + spreadsheet.getName());

        String[] splitFileAbsPath = spreadsheet.getAbsolutePath().split("\\.");
        if ((splitFileAbsPath[(splitFileAbsPath.length - 1)].equals("xls")) ||
                (splitFileAbsPath[(splitFileAbsPath.length - 1)].equals("xlsx")) ||
                (splitFileAbsPath[(splitFileAbsPath.length - 1)].equals("xlsm"))) {
            try {
                Workbook workbook = WorkbookFactory.create(new FileInputStream(spreadsheet));

                Sheet[] sheets = new Sheet[workbook.getNumberOfSheets()];
                FileOutputStream fileOut = new FileOutputStream(SpreadSheetWriter3.Directory.COMPLETED.toString() + "/" + spreadsheet.getName());
                for (int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++) {
                    sheets[sheetNum] = workbook.getSheetAt(sheetNum);
                }
                Iterator localIterator1;
                Row row;

                for (Sheet sheet : sheets) {
                    for (localIterator1 = sheet.iterator(); localIterator1.hasNext(); ) {
                        row = (Row) localIterator1.next();
                        for (Cell cell : row) {
                            try {
                                if (((!cell.getCellStyle().getLocked()) || cell.getCellType() != CellType.ERROR || cell.getCellType() != CellType.FORMULA)) {
                                    if (cell.getCellType() == CellType.NUMERIC && cell.getNumericCellValue() == 0) {
                                        if (cell.getNumericCellValue() == 0.0D) {
                                            if (sequential) {
                                                cell.setCellValue(seqInt);
                                            } else {
                                                cell.setCellValue(Helper.randInt(1, 10000));
                                            }
                                            numerictype++;
                                            seqInt++;
                                        }
                                    } else if (cell.getCellType() == CellType.STRING) {
                                        if (cell.getStringCellValue().equalsIgnoreCase("no data") || cell.getStringCellValue().isEmpty()) {
                                            blanktype++;
                                        } else if (cell.getStringCellValue().equalsIgnoreCase("item description") ||
                                                cell.getStringCellValue().startsWith("Enter") || cell.getStringCellValue().startsWith("enter") ||
                                                cell.getStringCellValue().equalsIgnoreCase("$0")) {
                                            if (sequential) {
                                                cell.setCellValue(seqInt);
                                            } else {
                                                cell.setCellValue(Helper.randInt(1, 10000));
                                            }
                                            stringtype++;
                                            seqInt++;
                                        }
                                    } else if (cell.getCellType() == CellType.BLANK) {
                                        /*if (sequential) {
                                            cell.setCellValue(seqInt);
                                        } else {
                                            cell.setCellValue(Helper.randInt(1, 10000));
                                        }*/
                                        blanktype++;
                                        //seqInt++;
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
                try {
                    FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                    evaluator.evaluateAll();
                } catch (Exception localException2) {
                }
                workbook.write(fileOut);
                fileOut.flush();
                fileOut.close();

                System.out.println("Completed: " + spreadsheet.getName());

                System.out.println("Numeric: " + numerictype);
                System.out.println("String: " + stringtype);
                System.out.println("Blank: " + blanktype);
            } catch (FileNotFoundException e) {
                System.out.println(e.getMessage());
            } catch (IOException e) {
                System.out.println(e.getMessage());
            }
        }
    }
}
