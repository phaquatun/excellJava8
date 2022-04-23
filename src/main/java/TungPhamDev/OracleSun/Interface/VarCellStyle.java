
package TungPhamDev.OracleSun.Interface;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Sheet;


public interface VarCellStyle {
    void varCellStyle (int numRowWorking , Sheet sheet , Cell numCellWorking ,Font font , CellStyle cellStyle);
}
