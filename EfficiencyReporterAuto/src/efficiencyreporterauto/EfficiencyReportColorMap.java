package efficiencyreporterauto;

import org.apache.poi.xssf.usermodel.IndexedColorMap;

/**
 *
 * @author JShepherd
 */
public class EfficiencyReportColorMap implements IndexedColorMap {

    @Override
    public byte[] getRGB(int index) {
        byte[] rgb = {(byte) 255, (byte) 255, (byte) 255};

        switch (index) {
            case 1:
                byte[] lightBlue = new byte[3];
                lightBlue[0] = (byte) 172;
                lightBlue[1] = (byte) 212;
                lightBlue[2] = (byte) 242;
                rgb = lightBlue;
                break;
            case 2:
                byte[] yellow = new byte[3];
                yellow[0] = (byte) 255;
                yellow[1] = (byte) 220;
                yellow[2] = (byte) 137;
                rgb = yellow;
            default:
                
        }

        return rgb;
    }
}
