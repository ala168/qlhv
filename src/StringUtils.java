import java.util.ArrayList;
import java.util.List;

public class StringUtils {

public static void main(String[] args) {
	List<String> tests = new ArrayList<String>();

	tests.add("Nguyen Khanh Linh SV12x6x11 thang 1 1 FT25331093320043   Ma giao dich   Trace680060 Trace 680060 				");
	tests.add("Nguyen Ba Gia Huy SV12x6x6 thang 11 				");
	tests.add("LE NGUYEN BAO CHAM SV12X5X3 THANG 1 1- Ma GD ACSP/ IE649116 				");
	tests.add("Hoang Minh Huy SV12x2x7 thang 11 FT  25331660462659   Ma giao dich  Trac  e789764 Trace 789764 				");
	tests.add("Mao Khanh An SV12x2x21 thang 11 				");
	tests.add("DUONG DINH LOI SV12X3X33 THANG 11-  Ma GD ACSP/ 3E346714 				");
	tests.add("Nguyen Viet Khoi SV12x3x32 thang 11     Ma giao dich  Trace502123 Trace   502123 				");
	tests.add("Truong Quoc Phap SV12x3x34 thang 11     Ma giao dich  Trace079591 Trace   079591 				");
	tests.add("Bui Hoang Gia Huy SV12x6x36 thang 1  1   Ma giao dich  Trace348190 Trace   348190 				");
	tests.add("Nguyen Tam Nhi SV12x1x23 thang 11     Ma giao dich  Trace186285 Trace 18  6285 				");
	tests.add("MBVCB.11895157500.829868.Luu gia Ba  o SV12x5x27 thang 11.CT tu 10282958  14 LUU THANH TRUNG toi 2588666868 C  ONG TY TNHH TRI THUC SAO VIET tai M B- Ma GD ACSP/ hr829868				");
	tests.add("NGUYEN QUANG DIEU SV12X2X37 THANG 1  1- Ma GD ACSP/ 4y088951 				");
	tests.add("Le Vuong Dat SV12x6x23 thang 11 FT2  5330382167215   Ma giao dich  Trace  932503 Trace 932503 				");
	tests.add("Cao Ha Trang SV12x3x8 thang 11 				");
	tests.add("Nguyen Tran Quang Vinh SV11x3x17 th  ang 11 				");
	tests.add("Nguyen Khoi Nguyen SV12x2x17 thang   11 				");
	tests.add("Ngo Dinh Nhat Minh SV12x5x24 thang   11 FT25330340696050   Ma giao dich    Trace233806 Trace 233806 				");
	tests.add("Nguyen Ngoc Phuc Khanh SV12x2x8 tha  ng 11 				");
	tests.add("MBVCB.11887061310.534603.Nguyen Nha  t Anh SV12x5x31 thang 11.CT tu 0011  004109521 PHUNG THI THANH THUY toi  2588666868 CONG TY TNHH TRI THUC SA O VIET tai MB- Ma GD ACSP/ ow534603				");
	tests.add("Nguyen Thanh Doan SV12x3x13 thang 1  1   Ma giao dich  Trace768472 Trace   768472 				");
	tests.add("VU QUANG SON SV12X2X13 THANG 11- Ma   GD ACSP/ Sb324799 				");
	tests.add("QR   Nguyen Dieu Anh SV12x5x1 thang   11- Ma GD ACSP/ 1y802560 				");
	tests.add("Do Phuong Thao Van SV12x6x18 thang   11 				");
	tests.add("QR   Khuc Ngoc Anh SV11x3x20 thang  11- Ma GD ACSP/ r1536331 				");
	tests.add("Pham Gia Khanh SV12x3x7 thang 11 				");
	tests.add("Vu Mai Phuong SV12x2x16 thang 11     Ma giao dich  Trace125444 Trace 125  444 				");
	tests.add("MBVCB.11882155338.244220.Vu Hai Yen   SV12x5x14 thang 11.CT tu 071100030  6741 VU KHAC HUNG toi 2588666868 CO  NG TY TNHH TRI THUC SAO VIET tai MB - Ma GD ACSP/ vk244220				");
	tests.add("Vu Hoang Viet SV12x3x36 thang 11 FT  25329487923642   Ma giao dich  Trac  e677368 Trace 677368 				");
	tests.add("QR   Nguyen Ngoc Son SV12x2x27 than  g 11- Ma GD ACSP/ uc300512 				");
	tests.add("DUONG THUY TIEN SV12X3X17 THANG 11-   Ma GD ACSP/ nT724018 				");
	tests.add("NGUYEN THANH BINH SV12x1x6 thang 11   chuyen tien FT25329673512500   Ma   giao dich  Trace307183 Trace 307183 				");
	tests.add("Bui Bao Ngoc SV12x1x22 thang 11 				");
	tests.add("MB 2588666868 Mai Anh Quan SV11x3x1  2 thang 11- Ma GD ACSP/ HB010270 				");
	tests.add("HOANG VIET ANH SV11X3X1 THANG 11- M  a GD ACSP/ 2Q448300 				");
	tests.add("Tran An Luong SV12x6x29 thang 11 				");
	tests.add("Nguyen Tan Dung SV12x1x18 thang 11   FT25329509495412   Ma giao dich  Tr  ace406195 Trace 406195 				");
	tests.add("Nguyen Hai Dang SV12x6x26 thang 11     Ma giao dich  Trace711328 Trace 7  11328 				");
	tests.add("MBVCB.11869978089.211313.Le Van Kie  n SV12x6x8 thang 11.CT tu 059100025  6643 LY THANH HUYEN toi 2588666868  CONG TY TNHH TRI THUC SAO VIET tai MB- Ma GD ACSP/ xn211313				");
	tests.add("Bui Quang Minh SV12x6x7 thang 11     Ma giao dich  Trace073119 Trace 073  119 				");
	tests.add("Dinh Thi Minh Phuong SV12x5x17 T11     Ma giao dich  Trace228449 Trace 2  28449 				");
	tests.add("Tran Thi Linh Phuong SV12x6x19 than  g 11 FT25328138222615   Ma giao dic  h  Trace860506 Trace 860506 				");
	tests.add("Nguyen Dang Cam Tu SV11x3x16 thang   11 FT25328520287008   Ma giao dich    Trace516055 Trace 516055 				");
	tests.add("Nguyen Xuan Phuc SV12x2x10 thang 11   FT25328742105204   Ma giao dich  T  race215966 Trace 215966 				");
	tests.add("Luu Minh Chau SV12x1x7 thang 11 				");
	tests.add("Nguyen Thi Kim Sa SV12x5x10 thang 1  1 				");
	tests.add("MBVCB.11860467554.037050.Do Thu Hie  n SV12 x1x8 TT HP T11 .CT tu 002100  0321578 MAI THI THANH QUYNH toi 258  8666868 CONG TY TNHH TRI THUC SAO V IET tai MB- Ma GD ACSP/ pb037050				");
	tests.add("CUSTOMER con Thu Quyen nop hp   Ma giao dich   Trace024237 Trace 024237");
	tests.add("SV12.1.3 con Thu Quyen nop hp   Ma giao dich   Trace024237 Trace 024237");

    System.out.println("===== BẮT ĐẦU TEST =====");

    for (String remark : tests) {
        RemarkInfo info = parseRemark(remark);
        System.out.println("----------------------------------");
        System.out.println(remark);
        System.out.println("Kết quả:");
        if (info == null) {
            System.out.println("Kết quả: NULL");
        } else {
            System.out.println("Kết quả: [" + info.mahv + "] Thang [" + info.month + "] Ten [" + info.tenhv + "]");
        }
    }
}


public static class RemarkInfo {
    public String tenhv;
    public String mahv;
    public int month=0;
}

public static RemarkInfo parseRemark(String remark1) {
    if (remark1 == null) return null;
    
    String remark = remark1.toUpperCase().replaceAll(" ", "");

    RemarkInfo info = new RemarkInfo();

    String regexMahv = "(SV\\d+[X\\.]\\d+[X\\.]\\d+)";
    try {
        java.util.regex.Pattern p = java.util.regex.Pattern.compile(regexMahv);
        java.util.regex.Matcher m = p.matcher(remark);

        if (!m.find()) {
            return null;
        }

        info.mahv = m.group(1).replace(".", "X");
    } catch (Exception ex) {
        return null;
    }
    
    int i=remark1.toUpperCase().indexOf("SV");
    if (i>0) info.tenhv=remark1.substring(0, i-1);
    
    String regexThang = "(THANG\\d+)";
    try {
        java.util.regex.Pattern p = java.util.regex.Pattern.compile(regexThang);
        java.util.regex.Matcher m = p.matcher(remark);

        if (m.find()) {
        	info.month = Integer.parseInt(m.group(1).replace("THANG", ""));
        	return info;
        }
    } catch (Exception ex) {}
    
    String regexThang2 = "(T\\d+)";
    try {
        java.util.regex.Pattern p = java.util.regex.Pattern.compile(regexThang2);
        java.util.regex.Matcher m = p.matcher(remark);

        if (m.find()) {
        	info.month = Integer.parseInt(m.group(1).replace("T", ""));
        }
    } catch (Exception ex) {}
    
    return info;
}

public static String extractAccountNumber(String text) {
    if (text == null) return "";

    // Tách thành từng dòng
    String[] lines = text.split("\\r?\\n");

    for (String line : lines) {
        // Tìm dòng có chứa "Account No"
        if (line.toLowerCase().contains("account no")) {

            // Tách theo dấu :
            int idx = line.indexOf(":");
            if (idx != -1 && idx < line.length() - 1) {
                String acc = line.substring(idx + 1).trim();
                return acc;
            }
        }
    }
    return "";
}


}
