import java.awt.Graphics2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.util.HashMap;
import java.util.Map;
import javax.imageio.ImageIO;

public class ZPLConverter {
    private int total;
    private int blackLimit = 380;
    private int widthBytes;
    private static final Map<Integer, String> mapCode = new HashMap<>();
    static {
        mapCode.put(1, "G");
        mapCode.put(2, "H");
        mapCode.put(3, "I");
        mapCode.put(4, "J");
        mapCode.put(5, "K");
        mapCode.put(6, "L");
        mapCode.put(7, "M");
        mapCode.put(8, "N");
        mapCode.put(9, "O");
        mapCode.put(10, "P");
        mapCode.put(11, "Q");
        mapCode.put(12, "R");
        mapCode.put(13, "S");
        mapCode.put(14, "T");
        mapCode.put(15, "U");
        mapCode.put(16, "V");
        mapCode.put(17, "W");
        mapCode.put(18, "X");
        mapCode.put(19, "Y");
        mapCode.put(20, "g");
        mapCode.put(40, "h");
        mapCode.put(60, "i");
        mapCode.put(80, "j");
        mapCode.put(100, "k");
        mapCode.put(120, "l");
        mapCode.put(140, "m");
        mapCode.put(160, "n");
        mapCode.put(180, "o");
        mapCode.put(200, "p");
        mapCode.put(220, "q");
        mapCode.put(240, "r");
        mapCode.put(260, "s");
        mapCode.put(280, "t");
        mapCode.put(300, "u");
        mapCode.put(320, "v");
        mapCode.put(340, "w");
        mapCode.put(360, "x");
        mapCode.put(380, "y");
        mapCode.put(400, "z");
    }

    public String convertfromImg(BufferedImage image) {
        String body = createBody(image);
        body = encodeHexAscii(body);
        return total + "," + total + "," + widthBytes + "," + body;
    }

    private String createBody(BufferedImage orginalImage) {
        StringBuilder sb = new StringBuilder();
        Graphics2D graphics = orginalImage.createGraphics();
        graphics.drawImage(orginalImage, 0, 0, null);
        int height = orginalImage.getHeight();
        int width = orginalImage.getWidth();
        int rgb, red, green, blue, index = 0;
        char[] auxBinaryChar = {'0', '0', '0', '0', '0', '0', '0', '0'};
        if (width % 8 > 0) {
            widthBytes = ((width / 8) + 1);
        } else {
            widthBytes = width / 8;
        }
        total=widthBytes*height;
        for (int h = 0; h < height; h++) {
            for (int w = 0; w < width; w++) {
                rgb = orginalImage.getRGB(w, h);
                red = (rgb >> 16) & 0x000000FF;
                green = (rgb >> 8) & 0x000000FF;
                blue = (rgb) & 0x000000FF;
                char auxChar = '1';
                int totalColor = red + green + blue;
                if (totalColor > blackLimit) {
                    auxChar = '0';
                }
                auxBinaryChar[index] = auxChar;
                index++;
                if (index == 8 || w == (width - 1)) {
                    sb.append(fourByteBinary(new String(auxBinaryChar)));
                    auxBinaryChar = new char[]{'0', '0', '0', '0', '0', '0', '0', '0'};
                    index = 0;
                }
            }
            sb.append("\n");
        }
        return sb.toString();
    }

    private String fourByteBinary(String binaryStr) {
        int decimal = Integer.parseInt(binaryStr, 2);
        if (decimal > 15) {
            return Integer.toString(decimal, 16).toUpperCase();
        } else {
            return "0" + Integer.toString(decimal, 16).toUpperCase();
        }
    }

    private String encodeHexAscii(String code) {
        int maxlinea = widthBytes * 2;
        StringBuilder sbCode = new StringBuilder();
        StringBuilder sbLinea = new StringBuilder();
        String previousLine = null;
        int counter = 1;
        char aux = code.charAt(0);
        boolean firstChar = false;
        for (int i = 1; i < code.length(); i++) {
            if (firstChar) {
                aux = code.charAt(i);
                firstChar = false;
                continue;
            }
            if (code.charAt(i) == '\n') {
                if (counter >= maxlinea && aux == '0') {
                    sbLinea.append(",");
                } else if (counter >= maxlinea && aux == 'F') {
                    sbLinea.append("!");
                } else if (counter > 20) {
                    int multi20 = (counter / 20) * 20;
                    int resto20 = (counter % 20);
                    sbLinea.append(mapCode.get(multi20));
                    if (resto20 != 0) {
                        sbLinea.append(mapCode.get(resto20)).append(aux);
                    } else {
                        sbLinea.append(aux);
                    }
                } else {
                    sbLinea.append(mapCode.get(counter)).append(aux);
                }
                counter = 1;
                firstChar = true;
                if (sbLinea.toString().equals(previousLine)) {
                    sbCode.append(":");
                } else {
                    sbCode.append(sbLinea);
                }
                previousLine = sbLinea.toString();
                sbLinea.setLength(0);
                continue;
            }
            if (aux == code.charAt(i)) {
                counter++;
            } else {
                if (counter > 20) {
                    int multi20 = (counter / 20) * 20;
                    int resto20 = (counter % 20);
                    sbLinea.append(mapCode.get(multi20));
                    if (resto20 != 0) {
                        sbLinea.append(mapCode.get(resto20)).append(aux);
                    } else {
                        sbLinea.append(aux);
                    }
                } else {
                    sbLinea.append(mapCode.get(counter)).append(aux);
                }
                counter = 1;
                aux = code.charAt(i);
            }
        }
        return sbCode.toString();
    }

    public void setBlacknessLimitPercentage(int percentage) {
        blackLimit = (percentage * 768 / 100);
    }

    public static void main(String[] args) throws Exception {
        int two;
        String one = args[0];
        if (args.length == 1) {
            two = 50;
        } else {
            two = Integer.parseInt(args[1]);
        }
        BufferedImage originalImage = ImageIO.read(new File(one));
        ZPLConverter zp = new ZPLConverter();
        zp.setBlacknessLimitPercentage(two);
        System.out.println(zp.convertfromImg(originalImage));
    }
}
