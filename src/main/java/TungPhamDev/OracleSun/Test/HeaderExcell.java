package TungPhamDev.OracleSun.Test;

import TungPhamDev.OracleSun.Interface.ExcellData;
import TungPhamDev.OracleSun.Interface.ReadExcell;

public class HeaderExcell {

    private String stt, name, address;

    public HeaderExcell() {
    }

    public String getStt() {
        return stt;
    }

    public void setStt(String stt) {
        this.stt = stt;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getAddress() {
        return address;
    }

    public void setAddress(String address) {
        this.address = address;
    }

    public String toString() {
        
        return stt + "\t" + name + "\t" + address;
    }

}
