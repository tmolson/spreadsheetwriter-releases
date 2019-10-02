package ca.on.gov.test.ssw;

import ca.on.gov.test.ssw.poi.ExcelWorker;
import ca.on.gov.test.ssw.util.Helper;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

public class SpreadSheetWriter3 {
    // Maximum number of concurrent excel files to read.
    private static int MAX_THREADS = 4;

    // Enumeration to hold the directory paths for spreadsheets
    public enum Directory {
        SOURCE("spreadsheets"),
        COMPLETED("completed-spreadsheets");

        private String directory;

        Directory(String directory) {
            this.directory = directory;
        }

        public String toString() {
            return directory;
        }
    }

    public static void main(String[] args)
    {
        boolean sequential;

        // Set up whether or not to write to the excel files sequentially
        if(args.length > 0)
            sequential = Boolean.parseBoolean(args[0]);
        else
            sequential = false;

        // Initialize the file directory paths (the folders for spreadsheets, and completed spreadsheets)
        initializeDirectories();

        // Set up a fixed thread pool of MAX_THREADS size
        ExecutorService pool = Executors.newFixedThreadPool(MAX_THREADS);

        File[] excelFiles = new File(Directory.SOURCE.toString()).listFiles();

        List<Runnable> excelWorkers = new ArrayList<>();
        for (File file : excelFiles)
        {
            excelWorkers.add(new ExcelWorker(file, sequential));

        }

        for(Runnable excelWorker : excelWorkers) {
            pool.execute(excelWorker);
        }

        pool.shutdown();
    }

    private static void initializeDirectories() {
        File completedDirectory = new File(Directory.COMPLETED.toString());
        File spreadsheetsDirectory = new File(Directory.SOURCE.toString());

        boolean validDirectoryStructure;

        validDirectoryStructure = Helper.existsOrCreateDirectory(completedDirectory);

        if(validDirectoryStructure)
            validDirectoryStructure = Helper.existsOrCreateDirectory(spreadsheetsDirectory);

        if(!validDirectoryStructure) {
            System.out.println("Missing spreadsheet directories '" + Directory.COMPLETED.toString() + "' and/or '" +
                    Directory.SOURCE.toString() + "' in root application directory");
            System.exit(-1);
        }
    }
}
