package org.example;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;

import org.apache.poi.ss.usermodel.*;

public class ExcelComparison {

    public static Map<String, List<String>> groupedCellChanges = new HashMap<>();
    public static Map<String, String> groupedDifferencesOnlyValues = new HashMap<>();
    public static String v1FolderPath = "";
    public static Set<String> sheetCollection = new HashSet<>();
    public static Set<Integer> rowChanges = new HashSet<>();
    public static Set<String> columnChanges = new HashSet<>();
    public static Map<String,String> rowChangesMessageMap = new HashMap<>();
    public static Map<String,String> columnChangesMessageMap = new HashMap<>();

    public static void main(String[] args) throws IOException {

         //Takes input of two folders from USER
        Scanner input = new Scanner(System.in);

        System.out.print("Enter path of version1 folder :: ");
        v1FolderPath = input.nextLine();

        System.out.print("Enter path of version2 folder :: ");
        String v2FolderPath = input.nextLine();

        System.out.print("Enter output folder path :: ");
        String outputFolderPath = input.nextLine();

        boolean areVersionsSame = compareExcelFilesInFolders(v1FolderPath, v2FolderPath,outputFolderPath);

        if (areVersionsSame) {
            System.out.println("Both version1 and version2 are the same.");
        }else{
            System.out.println("Versions are not same, difference stores inside folder diff");
        }
    }

    private static boolean compareExcelFilesInFolders(String v1FolderPath, String v2FolderPath,String outputFolderPath) throws IOException {
        File v1Folder = new File(v1FolderPath);
        File v2Folder = new File(v2FolderPath);

        // Check if the provided paths are directories
        if (!v1Folder.isDirectory() || !v2Folder.isDirectory()) {
            System.out.println("Invalid folder paths provided.");
            return false;
        }

        // Recursively compare Excel files in the provided directories
        boolean areVersionsSame = compareExcelFilesRecursively(v1Folder, v2Folder,outputFolderPath);

        return areVersionsSame;
    }
    private static boolean compareExcelFilesRecursively(File v1Folder, File v2Folder,String outputFolderPath) throws IOException {

        File[] v1Files = v1Folder.listFiles();
        File[] v2Files = v2Folder.listFiles();

        // Create subOutputFolderPath that mirrors the folder structure of the input
        String subOutputFolderPath = outputFolderPath + v1Folder.getAbsolutePath().substring(v1FolderPath.length());
        File subOutputFolder = new File(subOutputFolderPath);
        subOutputFolder.mkdirs();

        // Handles if files are not present in any of the folder.
        if (v1Files == null || v2Files == null) {
            System.out.println("One or Both folders are empty!!");
            return false;
        }

        // Sort files to ensure they are in the same order in both folders
        List<File> sortedV1Files = new ArrayList<>(Arrays.asList(v1Files));
        List<File> sortedV2Files = new ArrayList<>(Arrays.asList(v2Files));
        Collections.sort(sortedV1Files);
        Collections.sort(sortedV2Files);

        int numFiles = Math.min(sortedV1Files.size(), sortedV2Files.size());

        int changedFilesCount = 0;
        List<String> changedFileNames = new ArrayList<>();

        boolean folderHasChanges = false;

        for (int i = 0; i < numFiles; i++) {
            File v1File = sortedV1Files.get(i);
            File v2File = sortedV2Files.get(i);

            if (v1File.isDirectory() && v2File.isDirectory()) {
                // If both are directories, recursively compare them
                boolean areSubFoldersSame = compareExcelFilesRecursively(v1File, v2File, outputFolderPath);
                if (!areSubFoldersSame) {
                    changedFilesCount++;
                    changedFileNames.add(v1File.getName());
                }
            } else if (v1File.isFile() && v2File.isFile()) {
                // If both are files, compare them
                List<String> differences = compareExcelFileContents(v1File, v2File);
                if (!differences.isEmpty()) {
                    changedFilesCount++;
                    changedFileNames.add(v1File.getName());

                    // Create a file to store differences
                    String diffFileName = v1File.getName() + ".diff";
                    String diffFilePath = subOutputFolderPath + File.separator + diffFileName;
                    File diffFile = new File(diffFilePath);
                    FileWriter diffFileWriter = new FileWriter(diffFile);

                    // Write differences to the diff file
                    for (String difference : differences) {
                        diffFileWriter.write(difference + "\n");
                    }

                    diffFileWriter.close();
                }
                folderHasChanges = true;
            } else {
                // If one is a file and the other is a directory, they are not the same
                changedFilesCount++;
                changedFileNames.add(v1File.getName());
                folderHasChanges = true;
            }
        }

        // Compare file names in version1 and version2 to find added and removed files
        HashSet<String> version1FileNames = new HashSet<>();
        HashSet<String> version2FileNames = new HashSet<>();

        for (File file : sortedV1Files) {
            version1FileNames.add(file.getName());
        }

        for (File file : sortedV2Files) {
            version2FileNames.add(file.getName());
        }

        // Find added files present in version2 but not in version1
        HashSet<String> addedFiles = new HashSet<>(version2FileNames);
        addedFiles.removeAll(version1FileNames);

        // Find removed files present in version1 but not in version2
        HashSet<String> removedFiles = new HashSet<>(version1FileNames);
        removedFiles.removeAll(version2FileNames);

        if(folderHasChanges) {
            // Create the overall-summary file
            String overallSummaryFileName = "overall-summary.txt";
            String overallSummaryFilePath = subOutputFolderPath + File.separator + overallSummaryFileName;
            File overallSummaryFile = new File(overallSummaryFilePath);
            FileWriter overallSummaryFileWriter = new FileWriter(overallSummaryFile);

            // Write overall summary to the overall-summary file
            overallSummaryFileWriter.write("-----------------------------------------------------------------\n");
            overallSummaryFileWriter.write("Overall Summary of files in folder: " + v1Folder.getName() + "\n");
            overallSummaryFileWriter.write("-----------------------------------------------------------------\n");
            overallSummaryFileWriter.write("Total files compared: " + numFiles + "\n");
            overallSummaryFileWriter.write("Total changed files: " + changedFilesCount + "\n");
            overallSummaryFileWriter.write("Changed file names: "+"\n");
            for (String fileName : changedFileNames) {
                overallSummaryFileWriter.write("                    "+"* "+fileName+"\n");
            }
            overallSummaryFileWriter.write("\n");
            overallSummaryFileWriter.write("\n");

            if(!rowChanges.isEmpty() && !columnChanges.isEmpty()) {
                overallSummaryFileWriter.write("------------------------------------------------------\n");
                overallSummaryFileWriter.write("Global Changes in Files: " + "\n");
                overallSummaryFileWriter.write("------------------------------------------------------\n");
                overallSummaryFileWriter.write("\t\t\t\t* Row - ");
                for (Integer rowNumber : rowChanges) {
                    overallSummaryFileWriter.write(rowNumber + ",");
                }
                overallSummaryFileWriter.write("\n");
                overallSummaryFileWriter.write("\t\t\t\t* Column - ");
                for (String colNumber : columnChanges) {
                    overallSummaryFileWriter.write(colNumber + ",");
                }
                overallSummaryFileWriter.write("\n");
                overallSummaryFileWriter.write("\n");
            }

            if(!sheetCollection.isEmpty()) {
                overallSummaryFileWriter.write("------------------------------------------------------\n");
                overallSummaryFileWriter.write("Global Changes in Sheets: " + "\n");
                overallSummaryFileWriter.write("------------------------------------------------------\n");
                for (String sheetName : sheetCollection) {
                    overallSummaryFileWriter.write("                    " + "* " + sheetName + "\n");
                }
                overallSummaryFileWriter.write("\n");
                overallSummaryFileWriter.write("\n");
            }

            if(!rowChangesMessageMap.isEmpty()){
                overallSummaryFileWriter.write("------------------------------------------------------\n");
                overallSummaryFileWriter.write("Rows Affected: "+"\n");
                overallSummaryFileWriter.write("------------------------------------------------------\n");
                for(Map.Entry<String,String> entry : rowChangesMessageMap.entrySet())
                {
                    overallSummaryFileWriter.write("     * "+entry.getKey()+" ----> "+entry.getValue()+"\n");
                }
                overallSummaryFileWriter.write("\n");
                overallSummaryFileWriter.write("\n");
            }

            if(!columnChangesMessageMap.isEmpty()){
                overallSummaryFileWriter.write("------------------------------------------------------\n");
                overallSummaryFileWriter.write("Column Affected: "+"\n");
                overallSummaryFileWriter.write("------------------------------------------------------\n");
                for(Map.Entry<String,String> entry : columnChangesMessageMap.entrySet())
                {
                    overallSummaryFileWriter.write("     * "+entry.getKey()+" ----> "+entry.getValue()+"\n");
                }
                overallSummaryFileWriter.write("\n");
                overallSummaryFileWriter.write("\n");
            }


            if(!groupedCellChanges.isEmpty()) {
                overallSummaryFileWriter.write("------------------------------------------------------\n");
                overallSummaryFileWriter.write("Cell changes in Files" + "\n");
                overallSummaryFileWriter.write("------------------------------------------------------\n");
                overallSummaryFileWriter.write("Cell Number     :    FileName/SheetName" + "\n");
                for (Map.Entry<String, List<String>> entry : groupedCellChanges.entrySet()) {
                    String cellKey = entry.getKey();
                    List<String> filesAndSheets = entry.getValue();

                    // For extracting column number
                    String[] arr = cellKey.split("#");

                    int row = Integer.parseInt(arr[0]);
                    int column = Integer.parseInt(arr[1]);

                    // Here we get column in Excel type english character
                    char columnChar = (char) ('A' + (column - 1));

                    overallSummaryFileWriter.write("   " + columnChar + row + "     \t:\t" + "\n");

                    for (String fileAndSheet : filesAndSheets) {
                        String[] fileSheet = fileAndSheet.split("/");
                        String fName = fileSheet[0];
                        String sName = fileSheet[1];
                        String searchKey = fName + "#" + sName + "#" + cellKey;

                        overallSummaryFileWriter.write("\t\t\t" + "* " + fName + "/" + sName + " | " + groupedDifferencesOnlyValues.get(searchKey) + " \n");
                    }
                    overallSummaryFileWriter.write("\n");
                    overallSummaryFileWriter.write("\n");
                }
            }

            // Print added or removed files if there exist.
            if (!addedFiles.isEmpty() || !removedFiles.isEmpty()) {
                overallSummaryFileWriter.write("----------------------------------------------------------------------------------------\n");
                overallSummaryFileWriter.write("Files Added/Removed between " + v1Folder.getName() + " and " + v2Folder.getName() +"\n");
                overallSummaryFileWriter.write("----------------------------------------------------------------------------------------\n");
                if (!addedFiles.isEmpty()) {
                    overallSummaryFileWriter.write("Added files in " + v2Folder.getName() + ": " + "\n");
                    for (String fileName : addedFiles) {
                        overallSummaryFileWriter.write("                    "+"[+] "+fileName+"\n");
                    }
                }
                if (!removedFiles.isEmpty()) {
                    overallSummaryFileWriter.write("Removed files in " + v2Folder.getName() + "\n");
                    for (String fileName : removedFiles) {
                        overallSummaryFileWriter.write("                    "+"[-] "+fileName+"\n");
                    }
                }
            }

            //close
            overallSummaryFileWriter.close();
        }

        // Clear the groupedCellChanges,rowChanges,columnChanges,rowChangesMessageMap for the next folder comparison
        groupedCellChanges.clear();
        rowChanges.clear();
        columnChanges.clear();

        if(rowChangesMessageMap.isEmpty() && columnChangesMessageMap.isEmpty()){
            return changedFilesCount == 0 && addedFiles.isEmpty() && removedFiles.isEmpty();
        }
        rowChangesMessageMap.clear();
        columnChangesMessageMap.clear();
        return false;
    }

    // Function to check if two Excel files are equal
    private static List<String> compareExcelFileContents(File file1, File file2) throws IOException {
        List<String> differences = new ArrayList<>();

        Workbook workbook1 = null;
        Workbook workbook2 = null;

        // Map For grouped differences file wise
        Map<String, List<String>> groupedDifferences = null;

        try {
            workbook1 = WorkbookFactory.create(new FileInputStream(file1));
            workbook2 = WorkbookFactory.create(new FileInputStream(file2));

            // Check for number of sheets in both Excel files.
            int workbook1Sheets = workbook1.getNumberOfSheets();
            int workbook2Sheets = workbook2.getNumberOfSheets();

            if (workbook1Sheets != workbook2Sheets) {
                System.out.println("Sheets are not equal so excel file is not matched!!");
                System.exit(0);
            }

            // Check for sheet names
            List<String> workbook1SheetNames = new ArrayList<>();
            List<String> workbook2SheetNames = new ArrayList<>();

            for (int i = 0; i < workbook1Sheets; i++) {
                workbook1SheetNames.add(workbook1.getSheetName(i));
                workbook2SheetNames.add(workbook2.getSheetName(i));
            }

            Collections.sort(workbook1SheetNames);
            Collections.sort(workbook2SheetNames);

            if (!workbook1SheetNames.equals(workbook2SheetNames)) {
                System.out.println("Sheets Names are not matched in excel file");
                System.exit(0);
            }

            // Create a map to store grouped differences
            groupedDifferences = new HashMap<>();

            // Check for number of rows and number of cells in each row
            int totalSheet = workbook1Sheets;
            for (int i = 0; i < totalSheet; i++) {
                Sheet s1 = workbook1.getSheetAt(i);
                Sheet s2 = workbook2.getSheetAt(i);

                int numberOfRowsInSheet1 = s1.getLastRowNum() + 1;
                int numberOfRowsInSheet2 = s2.getLastRowNum() + 1;

                // Checking for number of rows in both sheets
                if (numberOfRowsInSheet1 != numberOfRowsInSheet2) {
                    int differenceInRows = numberOfRowsInSheet1 - numberOfRowsInSheet2;
                    String message;

                    if(differenceInRows > 0){
                        message = "[-] "+ Math.abs(differenceInRows) + " Rows Deleted!! ";
                    }else{
                        message = "[+] "+ Math.abs(differenceInRows) + " Rows Added!! ";
                    }
                    rowChangesMessageMap.put(file1.getName()+"/"+s1.getSheetName(),message);
                    continue;
                }

                // Checking for number of cells in each row of particular sheet
                for (int j = 0; j < numberOfRowsInSheet1; j++) {
                    Row row1 = s1.getRow(j);
                    Row row2 = s2.getRow(j);

                    // Handling null or Empty rows cases
                    if (row1 == null && row2 == null) {
                        continue; // Both rows are null or empty, move to the next row
                    } else if (row1 == null || row2 == null) {
                        System.out.println("Difference found at sheet " + s1.getSheetName() + ", row " + (j + 1) + " (One row is missing)");
                        continue;
                    }

                    int noOfColumnsInRow1 = row1.getLastCellNum();
                    int noOfColumnsInRow2 = row2.getLastCellNum();

                    // Checking for number of rows in both sheets
                    if (noOfColumnsInRow1 != noOfColumnsInRow2) {
                        int differenceInColumns = noOfColumnsInRow1 - noOfColumnsInRow2;
                        String message;

                        if(differenceInColumns > 0){
                            message = "[-] "+ Math.abs(differenceInColumns) + " Column Deleted!! ";
                        }else{
                            message = "[+] "+ Math.abs(differenceInColumns) + " Column Added!! ";
                        }
                        columnChangesMessageMap.put(file1.getName()+"/"+s1.getSheetName()+" -> [ row no.  -> "+(j+1)+" ]",message);
                        continue;
                    }

                    // Get the maximum number of cells to handle missing cells in a row
                    int maxCells = Math.max(row1.getLastCellNum(), row2.getLastCellNum());

                    for (int k = 0; k < maxCells; k++) {
                        Cell cell1 = row1.getCell(k);
                        Cell cell2 = row2.getCell(k);

                        // Helper method to get the formatted string representation of a cell value based on its data type
                        String cellValue1 = getFormattedCellValue(cell1);
                        String cellValue2 = getFormattedCellValue(cell2);

                        String cellValue1Type = "";
                        String cellValue2Type = "";

                        Integer i1 = tryParseInteger(cellValue1);
                        Float f1 = tryParseFloat(cellValue2);

                        Integer i2 = tryParseInteger(cellValue2);
                        Float f2 = tryParseFloat(cellValue1);

                        if (i1 != null) {
                            cellValue1Type = "Integer";
                        } else if (f1 != null) {
                            cellValue1Type = "Float";
                        }

                        if (i2 != null) {
                            cellValue2Type = "Integer";
                        } else if (f2 != null) {
                            cellValue2Type = "Float";
                        }

                        if (!cellValue1.equals(cellValue2)) {

                            // Create a cellKey which is helpful for storing (file name + sheet name + row number + column number)
                            String cellKey = file1.getName() + "#" + s1.getSheetName() + "#" + (j + 1) + "#" + (k + 1);
                            char columnChar = (char) ('A' + (k));

                            rowChanges.add(j+1);
                            columnChanges.add(columnChar+"");

                            sheetCollection.add(s1.getSheetName());

                            // Store the difference in the map, grouped by the cell key
                            List<String> diffnce = groupedDifferences.getOrDefault(cellKey, new ArrayList<>());

                            String groupedDifferenceCell1 = cellValue1;
                            String groupedDifferenceCell2 = cellValue2;

                            // Add the difference to the list
                            String difference = "Difference at Sheet: \"" + s1.getSheetName() + "\", Cell: "+ columnChar+ (j + 1) + ", "
                                    + "Original: " + cellValue1 +", "
                                    + "Revised: " + cellValue2 ;

                            if ((cellValue1Type.equals("Integer") && cellValue2Type.equals("Float"))) {
                                difference = "Difference at Sheet: \"" + s1.getSheetName() + "\", Cell: "+ columnChar+ (j + 1) + ", "
                                        + "Original: " + cellValue1 + " ( "+cellValue1Type+" )"+", "
                                        + "Revised: " + cellValue2 + " ( "+cellValue2Type+" ) ";
                                groupedDifferenceCell1 += " ( "+cellValue1Type+" )";
                                groupedDifferenceCell2 += " ( "+cellValue2Type+" )";
                            }else if((cellValue1Type.equals("Float") && cellValue2Type.equals("Integer"))){
                                difference = "Difference at Sheet: \"" + s1.getSheetName() + "\", Cell: "+ columnChar+ (j + 1) + ", "
                                        + "Original: " + cellValue1 + " ( "+cellValue1Type+" )"+", "
                                        + "Revised: " + cellValue2 + " ( "+cellValue2Type+" ) ";
                                groupedDifferenceCell1 += " ( "+cellValue1Type+" )";
                                groupedDifferenceCell2 += " ( "+cellValue2Type+" )";
                            }

                            groupedDifferencesOnlyValues.put(cellKey,"["+groupedDifferenceCell1+" -> "+groupedDifferenceCell2+"]");
                            differences.add(difference);
                            diffnce.add(difference);
                            groupedDifferences.put(cellKey,diffnce);
                        }
                    }
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (workbook1 != null) {
                workbook1.close();  //close workbook1.
            }
            if (workbook2 != null) {
                workbook2.close();  //close workbook2.
            }
        }

        // Print all the grouped differences together
        if (!groupedDifferences.isEmpty()) {
            for (Map.Entry<String, List<String>> entry : groupedDifferences.entrySet()) {
                //Extract key and get filename,sheetname and cell number.
                String overallKey = entry.getKey();
                String[] arr = overallKey.split("#");

                String key = arr[2] + "#" + arr[3];
                String fName = arr[0];
                String sName = arr[1];

                if (groupedCellChanges.get(key) == null) {
                    List<String> l = new ArrayList<>();
                    l.add(fName+ "/"+ sName);
                    groupedCellChanges.put(key, l);
                } else {
                    groupedCellChanges.get(key).add(fName+ "/"+ sName);
                }

            }
        }

        return differences; // Return the result
    }

    // Helper method to get the formatted string representation of a cell value based on its data type
    private static String getFormattedCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }

        switch (cell.getCellType()) {
            case STRING -> {
                return cell.getStringCellValue();
            }
            case NUMERIC -> {
                if (DateUtil.isCellDateFormatted(cell)) {
                    Date date = cell.getDateCellValue();
                    SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MMM-yyyy");
                    return dateFormat.format(date);
                }
                return String.valueOf(cell.getNumericCellValue());
            }
            case BOOLEAN -> {
                return String.valueOf(cell.getBooleanCellValue());
            }
            case FORMULA -> {
                return cell.getCellFormula();
            }
            case BLANK -> {
                return "";
            }
            default -> {
                return "Unsupported Cell Type";
            }
        }
    }

    public static Integer tryParseInteger(String value) {
        try {
            return Integer.parseInt(value);
        } catch (NumberFormatException e) {
            return null; // Parsing failed
        }
    }

    public static Float tryParseFloat(String value) {
        try {
            return Float.parseFloat(value);
        } catch (NumberFormatException e) {
            return null; // Parsing failed
        }
    }
}
