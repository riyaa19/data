import javax.swing.*;
import javax.swing.text.DefaultCaret;

import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class ExcelProcessingApp extends JFrame {
    private JTextArea resultTextArea;

    public ExcelProcessingApp() {
        setTitle("Excel Processing App");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(800, 600);
        setLayout(new BorderLayout());

        resultTextArea = new JTextArea();
        resultTextArea.setWrapStyleWord(true);
        resultTextArea.setLineWrap(true);
        JScrollPane scrollPane = new JScrollPane(resultTextArea);
        add(scrollPane, BorderLayout.CENTER);

        JPanel buttonPanel = new JPanel();
        add(buttonPanel, BorderLayout.SOUTH);

        JButton filterButton = new JButton("Filter Data");
        filterButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                filterData();
            }
        });
        buttonPanel.add(filterButton);

        JButton checkPSButton = new JButton("Check PS");
        checkPSButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                checkPS();
            }
        });
        buttonPanel.add(checkPSButton);

        JButton checkSOEButton = new JButton("Check SOE");
        checkSOEButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                checkSOE();
            }
        });
        buttonPanel.add(checkSOEButton);
        
        
        JButton checkBUButton = new JButton("Check BU");
        checkBUButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                checkBU();
            }
        });
        buttonPanel.add(checkBUButton);
        
        
        JButton checkOnOffButton = new JButton("Check On/Off");
        checkOnOffButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                checkOnOff();
            }
        });
        buttonPanel.add(checkOnOffButton);

        
        JButton checkProjectIDButton = new JButton("Check Project ID");
        checkProjectIDButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
            	checkProjectId();
            }
        });
        buttonPanel.add(checkProjectIDButton);
        
        
        JButton MissingInManPowerButton = new JButton("Check Missing in Manpower");
        MissingInManPowerButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
            	MissingInManpower();
            }
        });
        buttonPanel.add(MissingInManPowerButton);
        
        
        JButton MissingInBurnButton = new JButton("Check Missing in Burn");
        MissingInBurnButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
            	MissingInBurn();
            }
        });
        buttonPanel.add(MissingInBurnButton);
        // Add buttons for other checks (CheckBU, checkOnOff, CheckProjectId, MissingInBurn, MissingInManPower)
        DefaultCaret caret = (DefaultCaret) resultTextArea.getCaret();
        caret.setUpdatePolicy(DefaultCaret.ALWAYS_UPDATE);
        setVisible(true);
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(new Runnable() {
            @Override
            public void run() {
                new ExcelProcessingApp();
            }
        });
    }

    private void filterData() {
        JFileChooser fileChooser = new JFileChooser();
        int returnValue = fileChooser.showOpenDialog(null);
        if (returnValue == JFileChooser.APPROVE_OPTION) {
            File selectedFile = fileChooser.getSelectedFile();
            try {
                FileInputStream inputStream = new FileInputStream(selectedFile);
                XSSFWorkbook workbook = new XSSFWorkbook(inputStream);

            XSSFWorkbook newWorkbook = new XSSFWorkbook();

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                XSSFSheet sheet = workbook.getSheetAt(i);
                int rows = sheet.getPhysicalNumberOfRows();
                int cols = 0;

                for (int r = 0; r < rows; r++) {
                    XSSFRow row = sheet.getRow(r);
                    if (row != null) {
                        cols = Math.max(cols, row.getLastCellNum());
                    }
                }

                XSSFSheet newSheet = newWorkbook.createSheet(sheet.getSheetName());
                int newRowNumber = 0;
                Object[][] copyData = new Object[rows][cols];

                XSSFRow headerRow = sheet.getRow(0);
                if (headerRow != null) {
                    XSSFRow newHeaderRow = newSheet.createRow(newRowNumber++);
                    for (int c = 0; c < cols; c++) {
                        XSSFCell headerCell = headerRow.getCell(c);
                        if (headerCell != null) {
                            XSSFCell newHeaderCell = newHeaderRow.createCell(c);
                            newHeaderCell.setCellValue(headerCell.getStringCellValue());

                            XSSFCellStyle headerCellStyle = newWorkbook.createCellStyle();
                            headerCellStyle.cloneStyleFrom(headerCell.getCellStyle());
                            newHeaderCell.setCellStyle(headerCellStyle);
                        }
                    }
                }

                for (int r = 1; r < rows; r++) {
                    XSSFRow row = sheet.getRow(r);
                    if (row != null) {
                        XSSFCell yearCell = row.getCell(20);
                        XSSFCell monthCell = row.getCell(21);
                        XSSFCell onOffCell = row.getCell(12);

                        if (yearCell != null && yearCell.getCellType() == CellType.NUMERIC
                                && monthCell != null && monthCell.getCellType() == CellType.STRING
                                && yearCell.getNumericCellValue() == 2023
                                && monthCell.getStringCellValue().equalsIgnoreCase("October")) {
                            XSSFRow destRow = newSheet.createRow(newRowNumber++);
                            for (int c = 0; c < cols; c++) {
                                XSSFCell cell = row.getCell(c);
                                if (cell != null) {
                                    switch (cell.getCellType()) {
                                        case BOOLEAN:
                                            copyData[r][c] = cell.getBooleanCellValue();
                                            break;
                                        case STRING:
                                            copyData[r][c] = cell.getStringCellValue();
                                            break;
                                        case NUMERIC:
                                            copyData[r][c] = cell.getNumericCellValue();
                                            break;
                                        default:
                                            copyData[r][c] = null;
                                    }
                                    XSSFCell destCell = destRow.createCell(c);
                                    Object value = copyData[r][c];
                                    if (value != null) {
                                        if (value instanceof String) {
                                            if (c == 12) {
                                                String onOffValue = (String) value;
                                                if (onOffValue.equalsIgnoreCase("on")) {
                                                    destCell.setCellValue("onsite");
                                                } else if (onOffValue.equalsIgnoreCase("off")) {
                                                    destCell.setCellValue("offshore");
                                                }
                                            } else {
                                                destCell.setCellValue((String) value);
                                            }
                                        } else if (value instanceof Double) {
                                            destCell.setCellValue((Double) value);
                                        } else if (value instanceof Boolean) {
                                            destCell.setCellValue((Boolean) value);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            try (FileOutputStream outputStream = new FileOutputStream("./datafiles/filtered_data.xlsx")) {
                newWorkbook.write(outputStream);
                resultTextArea.append("The data has been filtered\n");
            }

            workbook.close();
            newWorkbook.close();
            inputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
            resultTextArea.append("An error occurred during data filtering\n");
        }
    } else {
        resultTextArea.append("File selection canceled\n");
    }
}

    private void checkPS() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setMultiSelectionEnabled(false);
        resultTextArea.append("Select the source file....\n");

        fileChooser.setDialogTitle("Select Source File");
        int returnValue = fileChooser.showOpenDialog(null);

        if (returnValue == JFileChooser.APPROVE_OPTION) {
            File sourceFile = fileChooser.getSelectedFile();
            resultTextArea.append("Select the destination file....\n");

            fileChooser.setDialogTitle("Select Destination File");
            returnValue = fileChooser.showOpenDialog(null);

            if (returnValue == JFileChooser.APPROVE_OPTION) {
                File destFile = fileChooser.getSelectedFile();

                try {
                    FileInputStream sourceFileInputStream = new FileInputStream(sourceFile);
                    FileInputStream destFileInputStream = new FileInputStream(destFile);

                    XSSFWorkbook sourceWorkbook = new XSSFWorkbook(sourceFileInputStream);
                    XSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0); // Assuming only one sheet in the source workbook

                    XSSFWorkbook destWorkbook = new XSSFWorkbook(destFileInputStream);
                    // Known column indices for "Name" and "Number"
                    int sourceNameColumnIndex = 2; // Assuming "Name" is in the third column (0-based index)
                    int sourceNumberColumnIndex = 0; // Assuming "Number" is in the first column
                    int destNameColumnIndex = 8; // Assuming "Name" is in the ninth column of destination sheets
                    int destNumberColumnIndex = 6; // Assuming "Number" is in the seventh column of destination sheets
            // Iterate through each sheet in the destination workbook
            for (int sheetIndex = 0; sheetIndex < destWorkbook.getNumberOfSheets(); sheetIndex++) {
                XSSFSheet destSheet = destWorkbook.getSheetAt(sheetIndex);

                // Iterate through the rows in the source workbook
                for (int i = 1; i <= sourceSheet.getLastRowNum(); i++) { // Start from 1 to skip the header row
                    XSSFRow sourceRow = sourceSheet.getRow(i);
                    if (sourceRow == null) continue; // Check for null row
                    XSSFCell sourceNameCell = sourceRow.getCell(sourceNameColumnIndex);
                    if (sourceNameCell == null) continue; // Check for null cell
                    String sourceName = sourceNameCell.getStringCellValue();

                    XSSFCell sourceNumberCell = sourceRow.getCell(sourceNumberColumnIndex);
                    if (sourceNumberCell == null) continue; // Check for null cell
                    double sourceNumber = getNumericValueFromCell(sourceNumberCell); // Validate the value

                    // Search for the same name in the current destination sheet
                    for (int j = 1; j <= destSheet.getLastRowNum(); j++) { // Start from 1 to skip the header row
                        XSSFRow destRow = destSheet.getRow(j);
                        if (destRow == null) continue; // Check for null row
                        XSSFCell destNameCell = destRow.getCell(destNameColumnIndex);
                        if (destNameCell == null) continue; // Check for null cell
                        String destName = destNameCell.getStringCellValue();

                        XSSFCell destNumberCell = destRow.getCell(destNumberColumnIndex);
                        if (destNumberCell == null) continue; // Check for null cell
                        double destNumber = getNumericValueFromCell(destNumberCell); // Validate the value

                        if (sourceName.equals(destName) && sourceNumber != destNumber) {
                            final String resultMessage = "PS number does not match for Name: " + sourceName + " in sheet: " + destSheet.getSheetName() + " in row " + (j + 1) + "\n";
                            SwingUtilities.invokeLater(new Runnable() {
                                @Override
                                public void run() {
                                    resultTextArea.append(resultMessage);
                                }
                            });
                            System.out.println(resultMessage);
                        }
                    }
                }
            }
            final String resultMessage = "PS check completed\n";
            SwingUtilities.invokeLater(new Runnable() {
                public void run() {
                    resultTextArea.append("PS check completed\n");
                }
            });

            sourceWorkbook.close();
            destWorkbook.close();
            sourceFileInputStream.close();
            destFileInputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
            resultTextArea.append("An error occurred during the PS check\n");
        }
    } else {
        resultTextArea.append("Destination file selection canceled\n");
    }
} else {
    resultTextArea.append("Source file selection canceled\n");
}
}

     private static double getNumericValueFromCell(XSSFCell cell) {
         double numericValue = 0.0;
         try {
             switch (cell.getCellType()) {
                 case NUMERIC:
                     numericValue = cell.getNumericCellValue();
                     break;
                 case STRING:
                     String cellValue = cell.getStringCellValue().trim();
                     if (cellValue.matches("-?\\d+(\\.\\d+)?")) {
                         numericValue = Double.parseDouble(cellValue);
                     }
                     break;
                 // Handle other cell types as needed
             }
         } catch (Exception ex) {
             ex.printStackTrace();
         }
         return numericValue;
     }
 

     private void checkSOE() {
    	    JFileChooser fileChooser = new JFileChooser();
    	    fileChooser.setMultiSelectionEnabled(false);
    	    fileChooser.setDialogTitle("Select Source File");
    	    int returnValue = fileChooser.showOpenDialog(null);

    	    if (returnValue == JFileChooser.APPROVE_OPTION) {
    	        File sourceFile = fileChooser.getSelectedFile();
    	        fileChooser.setDialogTitle("Select Destination File");
    	        returnValue = fileChooser.showOpenDialog(null);

    	        if (returnValue == JFileChooser.APPROVE_OPTION) {
    	            File destFile = fileChooser.getSelectedFile();
    	            try {
    	                FileInputStream sourceFileInputStream = new FileInputStream(sourceFile);
    	                FileInputStream destFileInputStream = new FileInputStream(destFile);

    	                XSSFWorkbook sourceWorkbook = new XSSFWorkbook(sourceFileInputStream);
    	                XSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0); // Assuming only one sheet in the source workbook

    	                XSSFWorkbook destWorkbook = new XSSFWorkbook(destFileInputStream);

            // Known column indices for "Name" and "Number"
            int sourceNameColumnIndex = 2; // Assuming "Name" is in the third column (0-based index)
            int sourceNumberColumnIndex = 1; // Assuming "Number" is in the second column
            int destNameColumnIndex = 8; // Assuming "Name" is in the ninth column of destination sheets
            int destNumberColumnIndex = 7; // Assuming "Number" is in the eighth column of destination sheets

            // Iterate through each sheet in the destination workbook
            for (int sheetIndex = 0; sheetIndex < destWorkbook.getNumberOfSheets(); sheetIndex++) {
                XSSFSheet destSheet = destWorkbook.getSheetAt(sheetIndex);

                // Iterate through the rows in the source workbook
                for (int i = 1; i <= sourceSheet.getLastRowNum(); i++) { // Start from 1 to skip the header row
                    XSSFRow sourceRow = sourceSheet.getRow(i);
                    if (sourceRow == null) continue; // Check for null row
                    XSSFCell sourceNameCell = sourceRow.getCell(sourceNameColumnIndex);
                    if (sourceNameCell == null) continue; // Check for null cell
                    String sourceName = sourceNameCell.getStringCellValue();

                    XSSFCell sourceNumberCell = sourceRow.getCell(sourceNumberColumnIndex);
                    if (sourceNumberCell == null) continue; // Check for null cell
                    String sourceNumber = getCellValueAsString(sourceNumberCell); // Handle error cells

                    // Search for the same name in the current destination sheet
                    for (int j = 1; j <= destSheet.getLastRowNum(); j++) { // Start from 1 to skip the header row
                        XSSFRow destRow = destSheet.getRow(j);
                        if (destRow == null) continue; // Check for null row
                        XSSFCell destNameCell = destRow.getCell(destNameColumnIndex);
                        if (destNameCell == null) continue; // Check for null cell
                        String destName = destNameCell.getStringCellValue();

                        XSSFCell destNumberCell = destRow.getCell(destNumberColumnIndex);
                        if (destNumberCell == null) continue; // Check for null cell
                        String destNumber = getCellValueAsString(destNumberCell); 
                        // Handle error cells
                        if(sourceName.equals(destName) ) {
                        //System.out.println(sourceName+" "+destName);
                        //System.out.println(sourceNumber+" "+destNumber);


                        if ( !sourceNumber.equalsIgnoreCase(destNumber) && !sourceNumber.equals("Error: #N/A")) {
                            System.out.println("SOE ID does not match for Name: " + sourceName + " in sheet: " + destSheet.getSheetName()+" in row number "+(j+1)+ "\n");
                            resultTextArea.append("SOE ID does not match for Name: " + sourceName + " in sheet: " + destSheet.getSheetName()+" in row number "+(j+1)+ "\n");
                            
                            
                        }
                    }
                    }
                }
                
            }
            resultTextArea.append("SOE check completed\n");


            // Close the workbooks
            sourceWorkbook.close();
            destWorkbook.close();
            sourceFileInputStream.close();
            destFileInputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
            resultTextArea.append("An error occurred during the SOE check\n");
        }
    } else {
        resultTextArea.append("Destination file selection canceled\n");
    }
} else {
    resultTextArea.append("Source file selection canceled\n");
}
}
    private static String getCellValueAsString(XSSFCell cell) {
        String cellValue = "";
        switch (cell.getCellType()) {
            case STRING:
                cellValue = cell.getStringCellValue();
                break;
            case NUMERIC:
                cellValue = String.valueOf(cell.getNumericCellValue());
                break;
            case ERROR:
                // Handle error cells here, for example:
                cellValue = "Error: " + cell.getErrorCellString();
                break;
            // Add more cases for other cell types if needed
        }
        return cellValue;
    }
 

    private void checkBU() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setMultiSelectionEnabled(false);
        fileChooser.setDialogTitle("Select Source File");
        int returnValue = fileChooser.showOpenDialog(null);

        if (returnValue == JFileChooser.APPROVE_OPTION) {
            File sourceFile = fileChooser.getSelectedFile();
            fileChooser.setDialogTitle("Select Destination File");
            returnValue = fileChooser.showOpenDialog(null);

            if (returnValue == JFileChooser.APPROVE_OPTION) {
                File destFile = fileChooser.getSelectedFile();

                try {
                    FileInputStream sourceFileInputStream = new FileInputStream(sourceFile);
                    FileInputStream destFileInputStream = new FileInputStream(destFile);

                    XSSFWorkbook sourceWorkbook = new XSSFWorkbook(sourceFileInputStream);
                    XSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0); // Assuming only one sheet in the source workbook

                    XSSFWorkbook destWorkbook = new XSSFWorkbook(destFileInputStream);


              // Known column indices for "Name" and "Number"
              int sourceNameColumnIndex = 7; // Assuming "BU" is in the seventh column (0-based index)
              int sourceNumberColumnIndex = 0; // Assuming "PS Number" is in the zeroth column
              int destNameColumnIndex = 2; // Assuming "BU" is in the second column of destination sheets
              int destNumberColumnIndex = 6; // Assuming "PS Number" is in the sixth column of destination sheets

              // Iterate through each sheet in the destination workbook
              for (int sheetIndex = 0; sheetIndex < destWorkbook.getNumberOfSheets(); sheetIndex++) {
                  XSSFSheet destSheet = destWorkbook.getSheetAt(sheetIndex);

                  // Iterate through the rows in the source workbook
                  for (int i = 1; i <= sourceSheet.getLastRowNum(); i++) { // Start from 1 to skip the header row
                      XSSFRow sourceRow = sourceSheet.getRow(i);
                      if (sourceRow == null) continue; // Check for null row
                      XSSFCell sourceNameCell = sourceRow.getCell(sourceNameColumnIndex);
                      if (sourceNameCell == null) continue; // Check for null cell
                      String sourceName = sourceNameCell.getStringCellValue();

                      XSSFCell sourceNumberCell = sourceRow.getCell(sourceNumberColumnIndex);
                      if (sourceNumberCell == null) continue; // Check for null cell
                      double sourceNumber = getNumericValueFromCell1(sourceNumberCell); // Validate the value

                      // Search for the same name in the current destination sheet
                      for (int j = 1; j <= destSheet.getLastRowNum(); j++) { // Start from 1 to skip the header row
                          XSSFRow destRow = destSheet.getRow(j);
                          if (destRow == null) continue; // Check for null row
                          XSSFCell destNameCell = destRow.getCell(destNameColumnIndex);
                          if (destNameCell == null) continue; // Check for null cell
                          String destName = destNameCell.getStringCellValue();

                          XSSFCell destNumberCell = destRow.getCell(destNumberColumnIndex);
                          if (destNumberCell == null) continue; // Check for null cell
                          double destNumber = getNumericValueFromCell1(destNumberCell); // Validate the value

                          if (!sourceName.equals(destName) && sourceNumber == destNumber) {
                              System.out.println("BU does not match for PS Number: " + sourceNumber + " in sheet: " + destSheet.getSheetName() +" in row "+(j+1 ));
                              resultTextArea.append("BU does not match for PS Number: " + sourceNumber + " in sheet: " + destSheet.getSheetName() +" in row "+(j+1 )+"\n");
                          }
                      }
                  }
              }

              // Close the workbooks
              sourceWorkbook.close();
              destWorkbook.close();
              sourceFileInputStream.close();
              destFileInputStream.close();
          } catch (Exception e) {
              e.printStackTrace();
              resultTextArea.append("An error occurred during the BU check\n");
          }
      } else {
          resultTextArea.append("Destination file selection canceled\n");
      }
  } else {
      resultTextArea.append("Source file selection canceled\n");
  }
}
      private static double getNumericValueFromCell1(XSSFCell cell) {
          double numericValue = 0.0;
          try {
              switch (cell.getCellType()) {
                  case NUMERIC:
                      numericValue = cell.getNumericCellValue();
                      break;
                  case STRING:
                      String cellValue = cell.getStringCellValue().trim();
                      if (cellValue.matches("-?\\d+(\\.\\d+)?")) {
                          numericValue = Double.parseDouble(cellValue);
                      }
                      break;
                  // Handle other cell types as needed
              }
          } catch (Exception ex) {
              ex.printStackTrace();
          }
          return numericValue;
      }
   
    
      private void checkOnOff() {
    	    JFileChooser fileChooser = new JFileChooser();
    	    fileChooser.setMultiSelectionEnabled(false);
    	    fileChooser.setDialogTitle("Select Source File");
    	    int returnValue = fileChooser.showOpenDialog(null);

    	    if (returnValue == JFileChooser.APPROVE_OPTION) {
    	        File sourceFile = fileChooser.getSelectedFile();
    	        fileChooser.setDialogTitle("Select Destination File");
    	        returnValue = fileChooser.showOpenDialog(null);

    	        if (returnValue == JFileChooser.APPROVE_OPTION) {
    	            File destFile = fileChooser.getSelectedFile();

    	            try {
    	                FileInputStream sourceFileInputStream = new FileInputStream(sourceFile);
    	                FileInputStream destFileInputStream = new FileInputStream(destFile);

    	                XSSFWorkbook sourceWorkbook = new XSSFWorkbook(sourceFileInputStream);
    	                XSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0); // Assuming only one sheet in the source workbook

    	                XSSFWorkbook destWorkbook = new XSSFWorkbook(destFileInputStream);
             // Known column indices for "Name" and "Number"
             int sourceNameColumnIndex = 2; // Assuming "Name" is in the third column (0-based index)
             int sourceLocationColumnIndex = 6; // Assuming "Number" is in the seventh column
                     
             int destNameColumnIndex = 8; // Assuming "Name" is in the ninth column of destination sheets
             int destLocationColumnIndex = 12; // Assuming "Location" is in the thirteenth column of destination sheets
             int a=0;

             // Iterate through each sheet in the destination workbook
             for (int sheetIndex = 0; sheetIndex < destWorkbook.getNumberOfSheets(); sheetIndex++) {
                 XSSFSheet destSheet = destWorkbook.getSheetAt(sheetIndex);

                 // Iterate through the rows in the source workbook
                 for (int i = 1; i <= sourceSheet.getLastRowNum(); i++) { // Start from 1 to skip the header row
                     XSSFRow sourceRow = sourceSheet.getRow(i);
                     if (sourceRow == null) continue; // Check for null row
                     XSSFCell sourceNameCell = sourceRow.getCell(sourceNameColumnIndex);
                     if (sourceNameCell == null) continue; // Check for null cell
                     String sourceName = sourceNameCell.getStringCellValue();

                     XSSFCell sourceLocationCell = sourceRow.getCell(sourceLocationColumnIndex);
                     if (sourceLocationCell == null) continue; // Check for null cell
                     String sourceLocation = getCellValueAsString1(sourceLocationCell); // Handle error cells

                     // Search for the same name in the current destination sheet
                     for (int j = 1; j <= destSheet.getLastRowNum(); j++) { // Start from 1 to skip the header row
                         XSSFRow destRow = destSheet.getRow(j);
                         if (destRow == null) continue; // Check for null row
                         XSSFCell destNameCell = destRow.getCell(destNameColumnIndex);
                         if (destNameCell == null) continue; // Check for null cell
                         String destName = destNameCell.getStringCellValue();

                         XSSFCell destLocationCell = destRow.getCell(destLocationColumnIndex);
                         if (destLocationCell == null) continue; // Check for null cell
                         String destLocation = getCellValueAsString1(destLocationCell); 
                         // Handle error cells
                         if(sourceName.equalsIgnoreCase(destName) ) {
                         //System.out.println(sourceName+" "+destName);
                         //System.out.println(sourceNumber+" "+destNumber);


                         if ( !sourceLocation.equalsIgnoreCase(destLocation) ) {
                             System.out.println("On/Off does not match for Name: " + sourceName + " in sheet: " + destSheet.getSheetName()+" in row number "+(j+1)+"\n");
                             resultTextArea.append("On/Off does not match for Name: " + sourceName + " in sheet: " + destSheet.getSheetName()+" in row number "+(j+1)+"\n");
                             a++;
                             
                             
                             
                         }
                     }
                     }
                 }
                 
             }
             resultTextArea.append("On/Off check completed\n");
            // System.out.println(a);

             // Close the workbooks
             sourceWorkbook.close();
             destWorkbook.close();
             sourceFileInputStream.close();
             destFileInputStream.close();
         } catch (Exception e) {
             e.printStackTrace();
             resultTextArea.append("An error occurred during the ON/OFF check\n");
         }
     } else {
         resultTextArea.append("Destination file selection canceled\n");
     }
 } else {
     resultTextArea.append("Source file selection canceled\n");
 }
}
     

     private static String getCellValueAsString1(XSSFCell cell) {
         String cellValue = "";
         switch (cell.getCellType()) {
             case STRING:
                 cellValue = cell.getStringCellValue();
                 break;
             case NUMERIC:
                 cellValue = String.valueOf(cell.getNumericCellValue());
                 break;
             case ERROR:
                 // Handle error cells here, for example:
                 cellValue = "Error: " + cell.getErrorCellString();
                 break;
             // Add more cases for other cell types if needed
         }
         return cellValue;
     }
  

    
    
     private void checkProjectId() {
    	    JFileChooser fileChooser = new JFileChooser();
    	    fileChooser.setMultiSelectionEnabled(false);
    	    fileChooser.setDialogTitle("Select Source File");
    	    int returnValue = fileChooser.showOpenDialog(null);

    	    if (returnValue == JFileChooser.APPROVE_OPTION) {
    	        File sourceFile = fileChooser.getSelectedFile();
    	        fileChooser.setDialogTitle("Select Destination File");
    	        returnValue = fileChooser.showOpenDialog(null);

    	        if (returnValue == JFileChooser.APPROVE_OPTION) {
    	            File destFile = fileChooser.getSelectedFile();

    	            try {
    	                FileInputStream sourceFileInputStream = new FileInputStream(sourceFile);
    	                FileInputStream destFileInputStream = new FileInputStream(destFile);

    	                XSSFWorkbook sourceWorkbook = new XSSFWorkbook(sourceFileInputStream);
    	                XSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0); // Assuming only one sheet in the source workbook

    	                XSSFWorkbook destWorkbook = new XSSFWorkbook(destFileInputStream);

              Map<String, String> sourceProjectData = new HashMap<>();

              for (int i = 1; i <= sourceSheet.getLastRowNum(); i++) {
                  XSSFRow sourceRow = sourceSheet.getRow(i);
                  if (sourceRow == null) continue;

                  XSSFCell nameCell = sourceRow.getCell(47); // Assuming project name is in the first column
                  XSSFCell idCell = sourceRow.getCell(12); // Assuming project ID is in the second column

                  if (nameCell != null && idCell != null) {
                      String projectName = nameCell.getStringCellValue().toLowerCase(); // Convert to lowercase
                      String projectID = idCell.getCellType() == CellType.STRING ?
                              idCell.getStringCellValue() :
                              String.valueOf((int) idCell.getNumericCellValue());
                      sourceProjectData.put(projectName, projectID);
                  }
              }

              for (int sheetIndex = 0; sheetIndex < destWorkbook.getNumberOfSheets(); sheetIndex++) {
                  XSSFSheet destSheet = destWorkbook.getSheetAt(sheetIndex);

                  for (int i = 1; i <= destSheet.getLastRowNum(); i++) {
                      XSSFRow destRow = destSheet.getRow(i);
                      if (destRow == null) continue;

                      XSSFCell nameCell = destRow.getCell(3); // Assuming project name is in the first column
                      XSSFCell idCell = destRow.getCell(1); // Assuming project ID is in the second column

                      if (nameCell != null && idCell != null) {
                          String projectName = nameCell.getStringCellValue().toLowerCase(); // Convert to lowercase
                          String projectID = idCell.getCellType() == CellType.STRING ?
                                  idCell.getStringCellValue() :
                                  String.valueOf((int) idCell.getNumericCellValue());

                          if (sourceProjectData.containsKey(projectName)) {
                              String expectedID = sourceProjectData.get(projectName);
                              if (!expectedID.equalsIgnoreCase(projectID)) {
                                  System.out.println("Mismatch for project: " + projectName +
                                          " in sheet: " + destSheet.getSheetName() + " in row: " + (i + 1)+ "\n");
                                  resultTextArea.append("Mismatch for project: " + projectName +
                                          " in sheet: " + destSheet.getSheetName() + " in row: " + (i + 1)+ "\n");
                              }
                          }
                      }
                  }
              }
           sourceWorkbook.close();
          destWorkbook.close();
          sourceFileInputStream.close();
          destFileInputStream.close();
      } catch (Exception e) {
          e.printStackTrace();
          resultTextArea.append("An error occurred during the Project ID check\n");
      }
  } else {
      resultTextArea.append("Destination file selection canceled\n");
  }
} else {
  resultTextArea.append("Source file selection canceled\n");
}
}

    
     private void MissingInManpower() {
    	    JFileChooser sourceFileChooser = new JFileChooser();
    	    sourceFileChooser.setDialogTitle("Select Source File");
    	    int sourceReturnValue = sourceFileChooser.showOpenDialog(null);

    	    if (sourceReturnValue == JFileChooser.APPROVE_OPTION) {
    	        File sourceFile = sourceFileChooser.getSelectedFile();

    	        JFileChooser destinationFileChooser = new JFileChooser();
    	        destinationFileChooser.setDialogTitle("Select Destination File");
    	        int destinationReturnValue = destinationFileChooser.showOpenDialog(null);

    	        if (destinationReturnValue == JFileChooser.APPROVE_OPTION) {
    	            File destFile = destinationFileChooser.getSelectedFile();

    	            try {
    	                FileInputStream sourceFileInputStream = new FileInputStream(sourceFile);
    	                FileInputStream destFileInputStream = new FileInputStream(destFile);

    	                XSSFWorkbook sourceWorkbook = new XSSFWorkbook(sourceFileInputStream);
    	                XSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0); // Assuming only one sheet in the source workbook

    	                XSSFWorkbook destWorkbook = new XSSFWorkbook(destFileInputStream);
    	                XSSFSheet destSheet = destWorkbook.getSheetAt(0); // Assuming only one sheet in the destination workbook

    	                Set<String> destData = new HashSet<>();

    	                // Store data from the destination sheet
    	                for (int i = 1; i <= destSheet.getLastRowNum(); i++) {
    	                    XSSFRow destRow = destSheet.getRow(i);
    	                    if (destRow == null) continue;

    	                    XSSFCell destCell = destRow.getCell(0); // Assuming data is in the first cell
    	                    if (destCell == null) continue;

    	                    String destValue = destCell.getStringCellValue();
    	                    destData.add(destValue);
    	                }

    	                // Compare source data with destination data
    	                for (int i = 1; i <= sourceSheet.getLastRowNum(); i++) {
    	                    XSSFRow sourceRow = sourceSheet.getRow(i);
    	                    if (sourceRow == null) continue;

    	                    XSSFCell sourceCell = sourceRow.getCell(7); // Assuming data is in the 8th cell (0-based index)
    	                    if (sourceCell == null) continue;

    	                    String sourceValue = sourceCell.getStringCellValue();

    	                    if (!destData.contains(sourceValue)) {
    	                        System.out.println("Data missing in manpower: " + sourceValue + "\n");
    	                        resultTextArea.append("Data missing in manpower: " + sourceValue + "\n");
    	                    }
    	                }

    	                sourceWorkbook.close();
    	                destWorkbook.close();
    	                sourceFileInputStream.close();
    	                destFileInputStream.close();
    	            } catch (Exception e) {
    	                e.printStackTrace();
    	                resultTextArea.append("An error occurred during the Missing in Manpower check\n");
    	            }
    	        } else {
    	            resultTextArea.append("Destination file selection canceled\n");
    	        }
    	    } else {
    	        resultTextArea.append("Source file selection canceled\n");
    	    }
    	}

    
     private void MissingInBurn() {
    	    JFileChooser sourceFileChooser = new JFileChooser();
    	    sourceFileChooser.setDialogTitle("Select Source File");
    	    int sourceReturnValue = sourceFileChooser.showOpenDialog(null);

    	    if (sourceReturnValue == JFileChooser.APPROVE_OPTION) {
    	        File sourceFile = sourceFileChooser.getSelectedFile();

    	        JFileChooser destinationFileChooser = new JFileChooser();
    	        destinationFileChooser.setDialogTitle("Select Destination File");
    	        int destinationReturnValue = destinationFileChooser.showOpenDialog(null);

    	        if (destinationReturnValue == JFileChooser.APPROVE_OPTION) {
    	            File destinationFile = destinationFileChooser.getSelectedFile();

    	            try {
    	                FileInputStream sourceFileInputStream = new FileInputStream(sourceFile);
    	                FileInputStream destinationFileInputStream = new FileInputStream(destinationFile);

    	                Workbook sourceWorkbook = new XSSFWorkbook(sourceFileInputStream);
    	                Workbook destinationWorkbook = new XSSFWorkbook(destinationFileInputStream);
            // Define the column numbers for the source and destination sheets
            int sourceColumnNumber = 2; // Column 2 in the source sheet
            int destinationColumnNumber = 8; // Column 8 in the destination sheets

            // Get the data from the source sheet
            Map<String, String> sourceData = getColumnData(sourceWorkbook.getSheetAt(0), sourceColumnNumber);

            // Initialize a set to store unique values from all destination sheets
            Set<String> destinationData = new HashSet<>();

            // Iterate through all sheets in the destination workbook
            for (int sheetIndex = 0; sheetIndex < destinationWorkbook.getNumberOfSheets(); sheetIndex++) {
                Sheet destinationSheet = destinationWorkbook.getSheetAt(sheetIndex);
                Map<String, String> sheetData = getColumnData(destinationSheet, destinationColumnNumber);
                destinationData.addAll(sheetData.values());
            }

            // Compare the data
            for (String key : sourceData.keySet()) {
                String sourceValue = sourceData.get(key);
                if (!destinationData.contains(sourceValue)) {
                    System.out.println("Data missing in burn: " + sourceValue+ "\n");
                    resultTextArea.append("Data missing in burn: " + sourceValue+ "\n");
                }
            }


            sourceWorkbook.close();
            destinationWorkbook.close();
            sourceFileInputStream.close();
            destinationFileInputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
            resultTextArea.append("An error occurred during Missing In Burn check\n");
        }
    } else {
        resultTextArea.append("Destination file selection canceled\n");
    }
} else {
    resultTextArea.append("Source file selection canceled\n");
}
}
    private static Map<String, String> getColumnData(Sheet sheet, int columnIndex) {
        Map<String, String> columnData = new HashMap<>();

        Iterator<Row> rowIterator = sheet.iterator();

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell cell = row.getCell(columnIndex);

            if (cell != null) {
                String cellValue = cell.getStringCellValue();
                columnData.put(Integer.toString(row.getRowNum() + 1), cellValue);
            }
        }

        return columnData;
    }
 }

    
    
    
    

