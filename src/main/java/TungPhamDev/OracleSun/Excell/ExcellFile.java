package TungPhamDev.OracleSun.Excell;

import TungPhamDev.OracleSun.Interface.ActionsExcell;
import TungPhamDev.OracleSun.Interface.ActionsUpdate;
import TungPhamDev.OracleSun.Interface.ExcellData;
import TungPhamDev.OracleSun.Interface.ReadExcell;
import TungPhamDev.OracleSun.Interface.SheetUpdate;
import TungPhamDev.OracleSun.Interface.WriteExcell;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import TungPhamDev.OracleSun.Interface.VarCellStyle;
import TungPhamDev.OracleSun.Interface.VarSheet;
import TungPhamDev.OracleSun.Interface.WriteExcellAsy;
import java.util.HashMap;
import java.util.Map;

public class ExcellFile {

    private String filePathExcell;

    private Workbook workbook = new XSSFWorkbook();
    private Sheet sheet;
    private Row numRow;
    private Cell numCell;
    private int numRowWoring;
    private InputStream inputStream;
    private OutputStream output;
    private VarCellStyle creater;

    public VarCellStyle getCreater() {
        return creater;
    }

    public void setCreater(VarCellStyle creater) {
        this.creater = creater;
    }

    public Sheet getSheet() {
        return sheet;
    }

    public Row getNumRow() {
        return numRow;
    }

    public Cell getNumCell() {
        return numCell;
    }

    public ExcellFile(String filePathExcell) {
        this.filePathExcell = filePathExcell;
    }

    /*
    ***
     */
    public ExcellFile readExcell(Sheet sheetWorking, ReadExcell readerExcell) {
        return call(() -> {
            methodReadExcell(sheetWorking, readerExcell);
        });
    }

    public ExcellFile readExcellFile(ReadExcell readerExcell) {
        Sheet sh = sheetWorkingAt(0);
        return call(() -> {
            methodOpendExcellToRead(() -> {
                methodReadExcell(sh, readerExcell);
            });
        });
    }

    public Sheet sheetWorkingName(String name) {
        try {
            inputStream = new FileInputStream(new File(filePathExcell));
            workbook = new XSSFWorkbook(inputStream);
            return workbook.getSheet(name);
        } catch (IOException ex) {
            ex.printStackTrace();
            return null;
        }
    }

    public Sheet sheetWorkingAt(int i) {
        try {
            inputStream = new FileInputStream(new File(filePathExcell));
            workbook = new XSSFWorkbook(inputStream);

            return workbook.getSheetAt(i);
        } catch (IOException ex) {
            ex.printStackTrace();
            return null;
        }

    }

    public void methodOpendExcellToRead(ActionsExcell excell) {
        methodOpendExcellToRead(1, excell);
    }

    public void methodOpendExcellToRead(Object obj, ActionsExcell excell) {
//        InputStream inputStream;
        try {
            inputStream = new FileInputStream(new File(filePathExcell));
            Workbook workbook = new XSSFWorkbook(inputStream);;

            if (obj instanceof String) {
                sheet = workbook.getSheet(String.valueOf(obj));
            }
            if (obj instanceof Integer) {
                sheet = workbook.getSheetAt((int) obj);
            }

            // Get all rows
            excell.runActions();

            closeReadExcell(workbook, inputStream);
        } catch (FileNotFoundException ex) {
            Logger.getLogger(ExcellFile.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(ExcellFile.class.getName()).log(Level.SEVERE, null, ex);
        }

    }

    public ExcellFile ReadExcellFromeTo(Sheet sheetWorking, int rowStart, int rowEnd, ReadExcell readerExcell) {
        return call(() -> {
            if (!checkFileExcellExis()) {
            } else {
                for (int i = rowStart; i < rowEnd; i++) {
                    numRow = sheetWorking.getRow(i);
                    readerExcell.headerObj();
                    Iterator<Cell> cellIterator = numRow.cellIterator();
                    while (cellIterator.hasNext()) {
                        numCell = cellIterator.next();
                        Object cellValue = getCellValue(numCell);
                        if (cellValue == null || cellValue.toString().isEmpty()) {
                            continue;
                        }
                        // Set value for book object
                        int columnIndex = numCell.getColumnIndex();
                        Object getCellValue = getCellValue(numCell);
                        readerExcell.setHeader(columnIndex, getCellValue);
                    }
                    readerExcell.readExcell(numRowWoring++);
                }
            }
        });
    }

    private void methodReadExcell(Sheet sheetWorking, ReadExcell readerExcell) {

        if (!checkFileExcellExis()) {
        } else {

//                    inputStream = new FileInputStream(new File(filePathExcell));
            // Get workbook
//                    workbook = new XSSFWorkbook(inputStream);
            // Get sheet sheet
//                    Sheet sheet = sheetWorking;
            // Get all rows
            Iterator<Row> iterator = sheetWorking.iterator();
            numRowWoring = 0;
            while (iterator.hasNext()) {
                numRow = iterator.next();
                if (numRow.getRowNum() == 0) {
                    // Ignore header
                    continue;
                }

                // object 
                readerExcell.headerObj();
                // Get all cells
                Iterator<Cell> cellIterator = numRow.cellIterator();
                // Read cells and set value for book object
                while (cellIterator.hasNext()) {
                    //Read cell
                    numCell = cellIterator.next();
                    Object cellValue = getCellValue(numCell);
                    if (cellValue == null || cellValue.toString().isEmpty()) {
                        continue;
                    }
                    // Set value for book object
                    int columnIndex = numCell.getColumnIndex();
                    Object getCellValue = getCellValue(numCell);
                    readerExcell.setHeader(columnIndex, getCellValue);
                }
                readerExcell.readExcell(numRowWoring++);
            }

//            closeReadExcell(workbook, inputStream);
        }
    }

    private Object getCellValue(Cell cells) {
        CellType cellType = cells.getCellType();
        Object cellValue = null;
        switch (cellType) {
            case FORMULA:
                workbook = cells.getSheet().getWorkbook();
                FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                cellValue = evaluator.evaluate(cells).getNumberValue();
                break;
            case NUMERIC:
                cellValue = cells.getNumericCellValue();
                break;
            case STRING:
                cellValue = cells.getStringCellValue();
                break;
            case _NONE:
                break;
            case BLANK:
                cellValue = null;
                break;
            case ERROR:
                break;

            default:
                break;
        }

        return cellValue;
    }

    /*
        ****
     */
    public ExcellFile writeExcellAsyn(String[] arrHeader, WriteExcellAsy writer) {
        return call(() -> {
            methodWriteExcellAsy(arrHeader, writer);
        });
    }

    public ExcellFile writeExcell(String[] arrHeader, WriteExcell writer) {
        // Haven't Style Cell
        return call(() -> {

            handlingWriteExcell(0, arrHeader, writer, (sheet, numRowWorking) -> {
            }, (numRowWorking, sheet, numCellWorking, font, cellStyle) -> {
            });

        });
    }

    public ExcellFile writeExcell(String[] arrHeader, WriteExcell writer, VarSheet varsheet) {
        return call(() -> {
            handlingWriteExcell(0, arrHeader, writer, varsheet, (numRowWorking, sheet, numCellWorking, font, cellStyle) -> {
            });

        });
    }

    public ExcellFile writeExcell(String[] arrHeader, WriteExcell writer, VarCellStyle createrStyle) {
        return call(() -> {
            handlingWriteExcell(0, arrHeader, writer, (sheet, numRowWorking) -> {
            }, createrStyle);

        });
    }

    public void methodCreateWriteExcell() {
        try {
            //Write the workbook in file system
//            File file = new File(filePathExcell);
//            if(file.exists()==false){
//                file.createNewFile();
//            }

            FileOutputStream out = new FileOutputStream(new File(filePathExcell));
            workbook.write(out);
            out.close();
            System.out.println(" written successfully on disk.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private ExcellFile handlingWriteExcell(int rowWorking, String[] arrHeader, WriteExcell writer, VarSheet varsheet, VarCellStyle createrStyle) {
        Font font = sheet.getWorkbook().createFont();
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        return call(() -> {

            List<Object[]> listArrData = new ArrayList<>();
            // add header excell
            if (arrHeader != null) {
                Object[] arrObjectHeader = new Object[arrHeader.length];
                for (int i = 0; i < arrObjectHeader.length; i++) {
                    arrObjectHeader[i] = arrHeader[i];
                }
                listArrData.add(arrObjectHeader);
            }

            // call back write here
            writer.writeExcell(listArrData);

            for (int i = 0; i < listArrData.size(); i++) {// create row , node i can generic
                Object[] arrObject = listArrData.get(i);
                numRow = sheet.createRow(i + rowWorking);
                numRowWoring = arrObject.length;

                for (int j = 0; j < numRowWoring; j++) {
                    Object valueCell = arrObject[j];
                    numCell = numRow.createCell(j);
                    numCell.setCellValue((String) valueCell);

                    //Cell Style
                    varsheet.varSheet(sheet, j);
                    createrStyle.varCellStyle(j, sheet, numCell, font, cellStyle);
                }
            }

//            System.out.println("heets tieenf");
            autosizeColumn(sheet, numRowWoring);
            methodCreateWriteExcell();

//            createExcell();
        });
    }

    private ExcellFile methodWriteExcellAsy(String[] arrHeader, WriteExcellAsy writer) {
        return call(() -> {

//            Object[] arrObjValue = new Object[arrHeader.length];
            Map<Integer, Object[]> mapData = new HashMap<>();
            if (arrHeader != null) {
                Object[] arrObjectHeader = new Object[arrHeader.length];
                for (int i = 0; i < arrObjectHeader.length; i++) {
                    arrObjectHeader[i] = arrHeader[i];
                }
                mapData.put(0, arrHeader);
            }

            writer.writeExcell(mapData);
            int rowWorking = 0;
            for (Integer integer : mapData.keySet()) {
                numRow = sheet.createRow(rowWorking++);
                Object[] arrObj = mapData.get(integer);
                numRowWoring = arrObj.length;
                int cellnum = 0;

                for (Object object : arrObj) {
                    numCell = numRow.createCell(cellnum++);
                    numCell.setCellValue((String) object);
                }
            }

            autosizeColumn(sheet, numRowWoring);
            methodCreateWriteExcell();

//            createExcell();
        });
    }

    public ExcellFile getSheet(int i) {
        return call(() -> {
            sheet = sheetWorkingAt(i);
        });
    }

    public ExcellFile getSheet(String strSheet) {
        return call(() -> {
            sheet = sheetWorkingName(strSheet);
        });
    }

    public ExcellFile createSheet() {
        return call(() -> {
            sheet = workbook.createSheet();

        });
    }

    public ExcellFile createSheet(String strSheet) {
        return call(() -> {
//            String safeName = WorkbookUtil.createSafeSheetName(strSheet);
            try {
                sheet = workbook.createSheet(strSheet);
            } catch (Exception e) {

            }

        });
    }

    private void autosizeColumn(Sheet sheet, int lastColumn) {
        for (int columnIndex = 0; columnIndex < lastColumn; columnIndex++) {
            sheet.autoSizeColumn(columnIndex);
        }
    }

    private ExcellFile createExcell() {
        return call(() -> {
            try {
                output = new FileOutputStream(filePathExcell);
                workbook.write(output);
                output.close();
            } catch (FileNotFoundException ex) {
                ex.printStackTrace();
            } catch (IOException ex) {
                ex.printStackTrace();
            }
        });
    }

    /*
    ****
     */
    public ExcellFile deleteSheet() {
        return call(() -> {
        });
    }

    public ExcellFile upDateValueCell(int row, int Colum, String value) {
        return call(() -> {
            methodOpendExcellToRead(() -> {
                Cell cell2Update = sheet.getRow(row).getCell(Colum);
                cell2Update.setCellValue(value);
            });
            methodCreateWriteExcell();
        });
    }

    public ExcellFile upDateValueRow(int rowLimit, int Colum, String value) {
        return call(() -> {
            methodOpendExcellToRead(() -> {
                for (int i = 0; i <= rowLimit; i++) {
                    for (int j = 0; j < Colum; j++) {
                        Cell cell2Update = sheet.getRow(i).getCell(Colum);
                        cell2Update.setCellValue(value);
                    }
                }

            });
            methodCreateWriteExcell();
        });
    }

    public ExcellFile upDateValueRow(Object obj, int rowLimit, int Colum, ActionsUpdate actionsUpdate) {
        return call(() -> {
            methodOpendExcellToRead(obj, () -> {
                for (int i = 1; i < rowLimit; i++) {
                    for (int j = 0; j < Colum; j++) {
                        Cell cell2Update = sheet.getRow(i).getCell(j);
                        actionsUpdate.update(cell2Update, i, j);
                    }
                }

            });
            methodCreateWriteExcell();
        });
    }

    public ExcellFile replaceValueCell(int row, int Colum, String replace, String value) {
        return call(() -> {
            methodOpendExcellToRead(() -> {
                for (int i = 0; i < row; i++) {
                    for (int j = 0; j < Colum; j++) {
                        Cell cell2Update = sheet.getRow(i).getCell(j);
                        Object cellValue = getCellValue(cell2Update);

                        String compare = replace.trim();
                        if (cellValue.toString().trim().contains(compare)) {
                            System.out.println(cellValue.toString());
                            cell2Update.setCellValue(value);
                        }
                    }
                }
                methodCreateWriteExcell();
            });
        });
    }

    public ExcellFile upDateValueRow(int row, WriteExcell write) {
        return call(() -> {
            handlingWriteExcell(row, null, write, (sheet, numRowWorking) -> {
            }, (numRowWorking, sheet, numCellWorking, font, cellStyle) -> {
            });

        });
    }

    /*
        ****
     */
    public boolean checkFileExcellExis() {
        File file = new File(filePathExcell);
        if (!file.exists()) {
            String[] arrNameFile = filePathExcell.split("\\\\");
            System.out.println(filePathExcell.split("\\\\")[arrNameFile.length - 1].split("\\.")[0]);
            return false;
        } else {
            return true;
        }
    }

    public void closeReadExcell(Workbook workbook, InputStream inputStream) {
        try {

            workbook.close();
            inputStream.close();

        } catch (IOException ex) {
            Logger.getLogger(ExcellFile.class
                    .getName()).log(Level.SEVERE, null, ex);
        }

    }

    private ExcellFile callStyle(int numRowWorking, Cell numCellWorking, Font font, CellStyle cellStyle, VarCellStyle creater) {

        return call(() -> {
//            System.out.println(numRowWorking);
            creater.varCellStyle(numRowWorking, sheet, numCell, font, cellStyle);

        });
    }

    public ExcellFile call(ExcellData data) {
        data.run();
        return this;
    }

    public static void main(String[] args) {

    }
}
