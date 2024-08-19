package gr.rtfm.util.xls2csv.commands;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintStream;
import java.math.BigDecimal;
import java.util.LinkedList;
import java.util.List;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;
import org.springframework.shell.standard.ShellComponent;
import org.springframework.shell.standard.ShellMethod;
import org.springframework.shell.standard.ShellOption;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

@ShellComponent
public class ConvertXLS {

    private static final org.apache.logging.log4j.Logger log = org.apache.logging.log4j.LogManager.getLogger();
    

    @ShellMethod("xls2csv")
    public void XLS2CSV(@ShellOption() String fromFile, @ShellOption(defaultValue = "stdout") String toFile) {

        log.info("fromFile:" + fromFile );

        FileInputStream is;
        try {
            is = new FileInputStream(fromFile);
            XSSFWorkbook workbook = new XSSFWorkbook(is);

            PrintStream outputStream = System.out;
            if ( !toFile.equals("stdout") )
            {
                outputStream = new PrintStream(toFile);
            }
            CSVPrinter printer = new CSVPrinter(outputStream, CSVFormat.DEFAULT);

            int numSheets = workbook.getNumberOfSheets();
            for (int s = 0; s < numSheets; s++) {
                String sheetName = workbook.getSheetName(s);
                sheetName = sheetName.trim();
                XSSFSheet sheet = workbook.getSheetAt(s);
                if (log.isInfoEnabled()) {
                    log.info("importXLS(): Sheet:" + s + " Name:'" + sheetName + "'");
                }

                dumpWorksheetToCSV(sheetName, sheet, printer);

            }

            workbook.close();
            printer.close();
        } catch (FileNotFoundException e) {
            log.error("Error opening toFile:"+toFile, e);
        } catch (IOException e) {
            log.error("IOException writing to toFile:"+toFile, e);
        }

    }

    private void dumpWorksheetToCSV(String sheetName, XSSFSheet sheet, CSVPrinter printer) {

        for (int r = sheet.getFirstRowNum(); r <= sheet.getLastRowNum(); r++) {
            List<String> rowData = new LinkedList<>();
            XSSFRow row = sheet.getRow(r);
            if ( row != null)
            {
                for (int c = row.getFirstCellNum(); c <= row.getLastCellNum(); c++) {
                    XSSFCell cell = row.getCell(c);
                    if (cell != null) {
                        rowData.add(getSafeString(cell));
                    } else {
                        rowData.add("");
                    }
                }    
            }
            try {
                printer.printRecord(rowData);
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    protected String getSafeString(XSSFCell cell) {
        if (cell == null) {
            return "";
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return BigDecimal.valueOf(cell.getNumericCellValue()).stripTrailingZeros().toPlainString();
            case FORMULA:
                CellType fType = cell.getCachedFormulaResultType();
                switch (fType) {
                    case STRING:
                        return cell.getStringCellValue();
                    case NUMERIC:
                        return BigDecimal.valueOf(cell.getNumericCellValue()).stripTrailingZeros().toPlainString();
                    default:
                        return cell.toString();
                }
            default:
                return cell.toString();
        }
    }
}