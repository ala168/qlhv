import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;

public class WordTemplate {

    public static void generateWordFromTemplate(String templatePath, String outputPath, Map<String, String> variables, String imagePath) throws IOException, InvalidFormatException {
        try (FileInputStream fis = new FileInputStream(templatePath);
             XWPFDocument document = new XWPFDocument(OPCPackage.open(fis))) {

            // Thay thế biến trong các đoạn văn bản
            for (XWPFParagraph p : document.getParagraphs()) {
                replaceVariablesInParagraph(p, variables, imagePath);
            }

            // Thay thế biến trong các bảng
            for (XWPFTable tbl : document.getTables()) {
                for (XWPFTableRow row : tbl.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph p : cell.getParagraphs()) {
                            replaceVariablesInParagraph(p, variables, imagePath);
                        }
                    }
                }
            }

            try (FileOutputStream out = new FileOutputStream(outputPath)) {
                document.write(out);
            }

        }
    }

    public static String removeControlCharacters(String input) {
        if (input == null) {
            return null;
        }
        // Regex: [\p{Cntrl}] bắt tất cả control characters
        return input.replaceAll("\\p{Cntrl}", "");
    }
    
    private static void replaceVariablesInParagraph(XWPFParagraph p, Map<String, String> variables, String imagePath) throws FileNotFoundException, IOException, InvalidFormatException {
        List<XWPFRun> runs = p.getRuns();
        if (runs != null) {
            for (XWPFRun r : runs) {
                String text = r.getText(0);
                
                //System.out.println("text=" + text);
                
                
                if (text != null) {
                    for (Map.Entry<String, String> entry : variables.entrySet()) {
                        String placeholder = "x"+ entry.getKey() + "x";
                        if (text.contains(placeholder)) {
                            text = text.replace(placeholder, entry.getValue());
                            r.setText(text, 0);
                            
                            //System.out.println("=> replace for key: " + entry.getKey());
                            
                        }
                    }
                    
                    if (!imagePath.isEmpty()){
	                    if (text != null && text.contains("ximagex")) {
	
	                        r.setText("", 0);
	
	
	                        try (InputStream is = new FileInputStream(imagePath)) {
	
	                            int format = getImageFormat(imagePath);
	                            if (format == -1) {
	                                System.err.println("Không thể xác định định dạng ảnh cho: " + imagePath);
	                                continue;
	                            }
	
	
	                            r.addPicture(is, format, imagePath, Units.toEMU(150), Units.toEMU(150));
	                        }
	                    }
	                }
                    
                }
            }
        }
    }    

    private static int getImageFormat(String imagePath) {
        String fileName = imagePath.toLowerCase();
        if (fileName.endsWith(".emf")) return XWPFDocument.PICTURE_TYPE_EMF;
        if (fileName.endsWith(".wmf")) return XWPFDocument.PICTURE_TYPE_WMF;
        if (fileName.endsWith(".pict")) return XWPFDocument.PICTURE_TYPE_PICT;
        if (fileName.endsWith(".jpeg") || fileName.endsWith(".jpg")) return XWPFDocument.PICTURE_TYPE_JPEG;
        if (fileName.endsWith(".png")) return XWPFDocument.PICTURE_TYPE_PNG;
        if (fileName.endsWith(".dib")) return XWPFDocument.PICTURE_TYPE_DIB;
        if (fileName.endsWith(".gif")) return XWPFDocument.PICTURE_TYPE_GIF;
        if (fileName.endsWith(".tiff")) return XWPFDocument.PICTURE_TYPE_TIFF;
        if (fileName.endsWith(".eps")) return XWPFDocument.PICTURE_TYPE_EPS;
        if (fileName.endsWith(".bmp")) return XWPFDocument.PICTURE_TYPE_BMP;
        if (fileName.endsWith(".wpg")) return XWPFDocument.PICTURE_TYPE_WPG;
        return -1;
    }

    public static void main(String[] args) {
        // Tạo một số dữ liệu giả định
        Map<String, String> variables = new java.util.HashMap<>();
        variables.put("thang", "09-2025");
        variables.put("tenhv", "Sản phẩm A");
        variables.put("tennh", "bidv");
        variables.put("sotk", "123456");
        variables.put("tentk", "vu khac tiep");
        variables.put("noidung", "Đây là mô tả chi tiết của sản phẩm A.");
        variables.put("loptt", "11A5");
        variables.put("math", "1");
        variables.put("van", "2");
        variables.put("eng", "3");
        variables.put("phy", "4");
        variables.put("chem", "5");
        variables.put("total", "6");
        variables.put("price", HVMain.formatCurrency("100000"));        
        variables.put("amt", HVMain.formatCurrency("20000000"));

        String templatePath = "10A2.docx";//"template.docx"; // Đảm bảo bạn có file template.docx
        String outputPath = "output.docx";
        String imagePath = "test.png"; // Đảm bảo bạn có file image.png BAO CAO\202509\qr\10A1.1.png

        try {
        	WordTemplate.generateWordFromTemplate(templatePath, outputPath, variables, imagePath);
            System.out.println("Tạo file Word thành công tại: " + outputPath);
            
            DocxMerger.openDesktop(outputPath);
        } catch (IOException | InvalidFormatException e) {
            e.printStackTrace();
        }
    }
}