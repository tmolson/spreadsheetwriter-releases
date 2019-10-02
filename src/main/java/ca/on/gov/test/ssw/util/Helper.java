package ca.on.gov.test.ssw.util;

import java.io.File;
import java.util.Random;

public class Helper {
    public static int randInt(int min, int max)
    {
        Random rand = new Random();

        int randomNum = rand.nextInt(max - min + 1) + min;

        return randomNum;
    }

    public static boolean existsOrCreateDirectory(File directory) {
        boolean result = true;

        if (!directory.exists()) {
            try {
                directory.mkdir();
            } catch (SecurityException e) {
                System.out.println(e.getMessage());
                result = false;
            }
        }

        return result;
    }
}
