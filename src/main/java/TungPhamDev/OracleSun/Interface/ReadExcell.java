package TungPhamDev.OracleSun.Interface;

public interface ReadExcell {

    void headerObj();

    void setHeader(int columnIndex, Object getCellValue);

    void readExcell(int rowWorking );
}
