package TungPhamDev.OracleSun.Test;

import TungPhamDev.OracleSun.Excell.ExcellFile;
import TungPhamDev.OracleSun.Interface.ExcellData;
import TungPhamDev.OracleSun.Interface.ReadExcell;
import TungPhamDev.OracleSun.Interface.SetValueCell;
import TungPhamDev.OracleSun.Interface.SheetWorking;
import TungPhamDev.OracleSun.Interface.TestIn;
import TungPhamDev.OracleSun.Interface.VarSheet;
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
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TestCase implements TestIn, VarSheet {
    
    private String filePathExcell;
    private static Workbook workbook = new XSSFWorkbook();
    private Sheet sheet;
    private Row numRow;
    private Cell numCell;
    private int numRowWoring;
    private InputStream inputStream;
    private OutputStream output;
    static List<HeaderExcell> listHeader = new ArrayList<>();
    
    public TestCase() {
    }
    
    public TestCase(String filePathExcell) {
        this.filePathExcell = filePathExcell;
    }
    
    private static HeaderExcell header;
    
    public static void main(String[] args) throws IOException {
//        List<HeaderExcell> listHeader = new ArrayList<>();
        String filePathExcell = "C:\\Users\\Admin\\Desktop\\FilterTotalData.xlsx";
        String sheetWorking = "checker 1";
        ExcellFile excell = new ExcellFile(filePathExcell);
        String[] arrHeader = {"stt", "name", "address"};
        
//        excell.ReadExcellFromeTo(excell.sheetWorkingAt(0), 100000, 200000, new ReadExcell() {
//            @Override
//            public void headerObj() {
//                 header = new HeaderExcell();
//            }
//
//            @Override
//            public void setHeader(int columnIndex, Object getCellValue) {
//               setValueCell(columnIndex, getCellValue);
//            }
//
//            @Override
//            public void readExcell(int rowWorking) {
//                System.out.println("row "+ rowWorking +"\t"+ header.toString());
//            }
//        });

//        testcase.createSheet(sheetWorking).writeExcell(arrHeader, (listArrData) -> {
//            listArrData.add(new Object[]{"human", "thật không đây", ":hay quá đi"});
//            listArrData.add(new Object[]{"khiing thèm chơi luôn ", null, "đã quá mà"});
//            listArrData.add(new Object[]{"kakaka được của nó", "chuẩn chỉ luôn ", null});
//            listArrData.add(new Object[]{"chết mẹ r", "chuẩn chỉ luôn ", "ối giời ơi"});
//            listArrData.add(new Object[]{"hay quas nhow ", "ddwuocj cua no", "chat choi nguoiw doi"});
//        }, (sheet, numRowWorking) -> System.out.println(numRowWorking));
////        
//        excell.createSheet().writeExcell(arrHeader, (listArrData) -> {
//            for (int i = 0; i < 10; i++) {
//                listArrData.add(new Object[]{
//                    "","","","","","",
//                });
//            }
//        });
        excell.createSheet().writeExcellAsyn(arrHeader, (mapData) -> {
            mapData.put(3, new Object[]{"được của nó ", "chất lượng ", "hay quá điu "});
            mapData.put(2, new Object[]{"được của nó 444", "chất lượng 444", "hay quá điu 444"});
            mapData.put(1, new Object[]{"được của nó 111", 111});
            
        });
//        
//        excell.upDateValueRow(3, 1, "human");
       

//        testcase.createSheet("huh");
//        testcase.readExcell(sheetWorking, header, setValueCell(), (data) -> System.out.println(data));
//        for (int i = 0; i < listHeader.size(); i++) {
//            HeaderExcell get = listHeader.get(i);
//            System.out.println(get.toString());
//        }
//        testcase.createSheet().writeExcell(arrHeader, (listArrData) -> { //istHeader.add((HeaderExcell)data)
//            listArrData.add(new Object[]{"human", "thật không đây", ":hay quá đi"});
//            listArrData.add(new Object[]{"khiing thèm chơi luôn ", null, "đã quá mà"});
//            listArrData.add(new Object[]{"kakaka được của nó", "chuẩn chỉ luôn ", null});
//        });
//        excell.readExcellFile(new ReadExcell() {
//            @Override
//            public void headerObj() {
//                header = new HeaderExcell();
//            }
//
//            @Override
//            public void setHeader(int columnIndex, Object getCellValue) {
//                setValueCell(columnIndex, getCellValue);
//            }
//
//            @Override
//            public void readExcell() {
//                System.out.println(header.toString());
//                listHeader.add(header);
//            }
//        });
//
//        excell.readExcell(excell.sheetWorkingAt(0), new ReadExcell() {
//            @Override
//            public void headerObj() {
//                header = new HeaderExcell();
//            }
//            
//            @Override
//            public void setHeader(int columnIndex, Object getCellValue) {
//                setValueCell(columnIndex, getCellValue);
//            }
//            
//            @Override
//            public void readExcell(int rowWorking) {
////                System.out.println(header.toString());
//                listHeader.add(header);
//            }
//        });
//        for (int i = 0; i < listHeader.size(); i++) {
//            HeaderExcell get = listHeader.get(i);
//            System.out.println(get.toString());
//        }
//        excell.createSheet().writeExcell(arrHeader, (listArrData) -> {
//            listArrData.add(new Object[]{"woman222", "thật không đây", ":hay quá đi"});
//            listArrData.add(new Object[]{"khiing thèm chơi luôn ", null, "đã quá mà"});
//        });
//        excell.getSheet("sheetExport").upDateValueRow(1, (listArrData) -> {
//            listArrData.add(new Object[]{
//                "đây này", "được cuatr nó", "đây là update"
//            });
//        });
    }
    
    private static void setValueCell(int columnIndex, Object getCellValue) {
        switch (columnIndex) {
            case 0:
                header.setStt(getCellValue.toString());
                break;
            case 1:
                header.setName(getCellValue.toString());
                break;
            case 2:
                header.setAddress(getCellValue.toString());
                break;
        }
        
    }
    
    public TestCase readExcell(Sheet sheet, ReadExcell readerExcell) {
        if (!checkFileExis()) {
        } else {
            // Get file
            try {
                inputStream = new FileInputStream(new File(filePathExcell));
                // Get workbook
                workbook = workbook = new XSSFWorkbook(inputStream);

                // Get sheet sheet
                // Get all rows
                Iterator<Row> iterator = sheet.iterator();
                int rowWorking = 0;                
                
                while (iterator.hasNext()) {
                    numRow = iterator.next();
                    if (numRow.getRowNum() == 0) {
                        // Ignore header
                        continue;
                    }

                    // object 
                    readerExcell.headerObj();
//                    };
//                    readExcell.headerObj();
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
                    readerExcell.readExcell(rowWorking++);
                    
                }
                
                workbook.close();
                inputStream.close();
                
            } catch (FileNotFoundException ex) {
                ex.printStackTrace();
            } catch (IOException ex) {
                ex.printStackTrace();
            }
            
        }
        
        return call();
        
    }
    
    public Sheet sheetWorkingName(String name) {
        return sheet = workbook.getSheet(name);
    }
    
    public Sheet sheetWorkingAt(int i) {
        
        return sheet = workbook.getSheetAt(i);
    }
    
    public Object getCellValue(Cell cells) {
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
    
    public TestCase writeExcell(String[] arrHeader, WriteExcell writer) throws IOException {
        List<Object[]> listArrData = new ArrayList<>();
        // add header excell
        if (arrHeader != null) {
            Object[] arrObjectHeader = new Object[arrHeader.length];
            for (int i = 0; i < arrObjectHeader.length; i++) {
                arrObjectHeader[i] = arrHeader[i];
            }
            listArrData.add(arrObjectHeader);
        }
        
        writer.writeExcell(listArrData);
        
        for (int i = 0; i < listArrData.size(); i++) {// create row , node i can generic
            Object[] arrObject = listArrData.get(i);
            numRow = sheet.createRow(i);
            numRowWoring = arrObject.length;
            
            for (int j = 0; j < numRowWoring; j++) {
                Object valueCell = arrObject[j];
                numCell = numRow.createCell(j);
                numCell.setCellValue((String) valueCell);
                
            }
            
        }
        
        autosizeColumn(sheet, numRowWoring);
        createExcell();
        
        return call();
    }
    
    public void autosizeColumn(Sheet sheet, int lastColumn) {
        for (int columnIndex = 0; columnIndex < lastColumn; columnIndex++) {
            sheet.autoSizeColumn(columnIndex);
        }
    }
    
    public TestCase createSheet() {
        sheet = workbook.createSheet();
        
        return call();
    }
    
    public TestCase createSheet(String strSheet) {
        
        String safeName = WorkbookUtil.createSafeSheetName(strSheet);
        sheet = workbook.createSheet(safeName);
        return call();
    }
    
    public TestCase createExcell() throws FileNotFoundException, IOException {
        output = new FileOutputStream(filePathExcell);
        workbook.write(output);
        
        output.close();
//        workbook.close();
        return call();
    }
    
    private boolean checkFileExis() {
        File file = new File(filePathExcell);
        if (!file.exists()) {
            String[] arrNameFile = filePathExcell.split("\\\\");
            System.out.println(filePathExcell.split("\\\\")[arrNameFile.length - 1].split("\\.")[0]);
            return false;
        } else {
            return true;
        }
    }
    
    private TestCase call() {
        return this;
    }
    
    @Override
    public void test() {
        System.out.println("affeter");
    }
    
    @Override
    public void varSheet(Sheet sheet, int numRowWorking) {
        
        System.out.println("xử lý style tại đây ");
        
    }
    
}
