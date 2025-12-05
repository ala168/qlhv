import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException; // Cần thiết cho việc merge DOCX
import org.openxmlformats.schemas.drawingml.x2006.main.CTGraphicalObject;
import org.openxmlformats.schemas.drawingml.x2006.main.CTGraphicalObjectData;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody; // Cần thiết cho việc merge DOCX

import java.io.File; // Cần thiết cho Files.exists
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URLEncoder;
import java.nio.file.Files; // Java NIO cho đường dẫn
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.*; // Cho SQLite
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream; // Cho Files.walk
import java.awt.Desktop; // Để mở file

import javax.xml.namespace.QName;

public class WordGenerator {

    private static final String QR_IMAGE_BASE_DIR = "baseDir/qr/"; // Thay đổi nếu đường dẫn khác
    private static final int IMAGE_WIDTH_EMU = Units.toEMU(150); // Chiều rộng ảnh QR
    private static final int IMAGE_HEIGHT_EMU = Units.toEMU(150); // Chiều cao ảnh QR

    // Hàm chính để tạo tài liệu tổng hợp từ template và dữ liệu SQLite
    public static void generateMergedWordFromSqlite(String templatePath, String outputMergedPath, String baseQrDir) throws Exception {
        
    	
    	
    	 // 1. Tạo thư mục
        File rootDir = new File("BAO CAO");
        if (!rootDir.exists()) rootDir.mkdirs();

        File baseDir = new File(rootDir, String.valueOf(202509));
        if (!baseDir.exists()) baseDir.mkdirs();

        File qrDir = new File(baseDir, "qr");
        if (qrDir.exists()) {
        	qrDir.delete();
        }
        qrDir.mkdirs();

        File hvDir = new File(baseDir, "hv");
        if (hvDir.exists()) {
        	hvDir.delete();
        }
        hvDir.mkdirs();  	
    	
    	
    	
    	XWPFDocument mergedDocument = null; // Tài liệu đích cuối cùng

        try (Connection conn = DriverManager.getConnection(HVMain.URL_DB);
             Statement stmt = conn.createStatement();
             ResultSet rs = stmt.executeQuery("SELECT id, thang, tenhv, mahv, loptt, monToan, monVan, monAnh, monLy, monHoa, sotk, tentk, tennh, dongia, tongsongay FROM tonghopthang where loptt='11A6'")) {

            int recordCount = 0;
            while (rs.next()) {
                recordCount++;
                // 1. Chuẩn bị dữ liệu biến cho template
                Map<String, String> variables = new HashMap<>();
                
                variables.put("thang", rs.getString("thang"));
                variables.put("tenhv", rs.getString("tenhv"));
                variables.put("mahv", rs.getString("mahv"));
                variables.put("loptt", rs.getString("loptt"));
                
                variables.put("math", rs.getString("monToan"));
                variables.put("van", rs.getString("monVan"));
                variables.put("eng", rs.getString("monAnh"));
                variables.put("phy", rs.getString("monLy"));
                variables.put("chem", rs.getString("monHoa"));
                variables.put("sotk", rs.getString("sotk"));
                variables.put("tentk", rs.getString("tentk"));
                variables.put("tennh", rs.getString("tennh"));
                variables.put("price", rs.getString("dongia"));
                variables.put("total", rs.getString("tongsongay"));
                
                
                // Add any other variables you need, e.g., computed ones
                // variables.put("total", calculateTotal(rs));
                // variables.put("price", formatCurrency(rs.getDouble("dongia")));
                // variables.put("amt", formatCurrency(rs.getDouble("tongsongay") * rs.getDouble("dongia")));
                
                // 2. Xác định đường dẫn ảnh QR
                String mahv = rs.getString("mahv");
                //File imagePath = new File("./BAOCAO/202509/qr/",mahv+".png");
                //String qrImagePath = baseQrDir +"/"+ mahv + ".png";
                
                
                double dongia = rs.getDouble("dongia");
                int tongsongay = rs.getInt("tongsongay");
                double tongtien = dongia * tongsongay;

                variables.put("amt", String.valueOf(tongtien));
                
                String des = mahv + " " + rs.getString("tenhv") + " TT ";
                String encodedDes = URLEncoder.encode(HVMain.removeVietnameseAccents(des), java.nio.charset.StandardCharsets.UTF_8.name());

                String url = "https://qr.sepay.vn/img?bank=MB&acc=" + rs.getString("sotk") +
                        "&template=qronly&amount=" + String.valueOf(tongtien) +
                        "&download=DOWNLOAD&des=" + encodedDes;

                //System.out.println(url);

                String qrImagePath = HVMain.downloadQrCodeWithExt(url, qrDir.getAbsolutePath(), mahv);
                System.out.println(qrImagePath);                
                
                // 3. Tạo một tài liệu tạm thời từ template và thay thế biến/ảnh
                XWPFDocument tempDoc = createDocumentFromTemplate(templatePath, variables, qrImagePath);

                // 4. Merge tài liệu tạm thời vào tài liệu chính
                if (mergedDocument == null) {
                    // Lần đầu tiên, mergedDocument chính là tempDoc
                    mergedDocument = tempDoc;
                } else {
                    // Thêm ngắt trang trước khi thêm nội dung mới
                    XWPFParagraph pageBreakParagraph = mergedDocument.createParagraph();
                    XWPFRun pageBreakRun = pageBreakParagraph.createRun();
                    pageBreakRun.addBreak(BreakType.PAGE);

                    // Map old image IDs to new image IDs for the current document
                    Map<String, String> oldToNewImageIds = copyImagesFromDocToDoc(tempDoc, mergedDocument);

                    CTBody nextDocBody = tempDoc.getDocument().getBody();

                    for (int k = 0; k < nextDocBody.sizeOfPArray(); k++) {
                        XWPFParagraph newPara = mergedDocument.createParagraph();
                        newPara.getCTP().set(nextDocBody.getPArray(k));
                        updateImageReferencesInParagraph(newPara, oldToNewImageIds);
                    }

                    for (int k = 0; k < nextDocBody.sizeOfTblArray(); k++) {
                        XWPFTable newTable = mergedDocument.createTable();
                        newTable.getCTTbl().set(nextDocBody.getTblArray(k));
                        updateImageReferencesInTable(newTable, oldToNewImageIds);
                    }
                    tempDoc.close(); // Đóng tài liệu tạm thời sau khi merge
                }
            }

            if (recordCount == 0) {
                System.out.println("No data found in the database. Creating an empty document or template if needed.");
                // Tùy chọn: Nếu không có dữ liệu, tạo một tài liệu trống hoặc một bản sao template
                if (mergedDocument == null) { // Nếu chưa có tài liệu nào được tạo
                     try (FileInputStream fis = new FileInputStream(templatePath)) {
                        mergedDocument = new XWPFDocument(OPCPackage.open(fis));
                    }
                    System.out.println("Created a document from template as no data was found.");
                }
            }
            
            // 5. Ghi tài liệu chính ra file
            if (mergedDocument != null) {
                try (FileOutputStream out = new FileOutputStream(outputMergedPath)) {
                    mergedDocument.write(out);
                }
            } else {
                System.err.println("Merged document is null, no output file created.");
            }

        } finally {
            if (mergedDocument != null) {
                try {
                    mergedDocument.close();
                } catch (IOException e) {
                    System.err.println("Error closing merged document: " + e.getMessage());
                }
            }
        }
    }

    // --- Helper function to create a single document from template with replacements ---
    private static XWPFDocument createDocumentFromTemplate(String templatePath, Map<String, String> variables, String qrImagePath) throws IOException, InvalidFormatException {
        XWPFDocument document;
        try (FileInputStream fis = new FileInputStream(templatePath)) {
            document = new XWPFDocument(OPCPackage.open(fis));
        }

        // Thay thế biến và ảnh trong các đoạn văn bản
        for (XWPFParagraph p : document.getParagraphs()) {
            replaceVariablesAndImageInParagraph(p, variables, qrImagePath);
        }

        // Thay thế biến và ảnh trong các bảng
        for (XWPFTable tbl : document.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        replaceVariablesAndImageInParagraph(p, variables, qrImagePath);
                    }
                }
            }
        }
        return document;
    }


    // --- Modified replaceVariablesInParagraph to handle images and cleanup ---
    private static void replaceVariablesAndImageInParagraph(XWPFParagraph p, Map<String, String> variables, String qrImagePath) throws FileNotFoundException, IOException, InvalidFormatException {
        boolean imageReplaced = false; // Cờ để theo dõi xem ảnh đã được thay thế chưa

        for (int i = 0; i < p.getRuns().size(); i++) {
            XWPFRun r = p.getRuns().get(i);
            String text = (r.getText(0) == null) ? "" : r.getText(0);
            
            // Thay thế biến
            for (Map.Entry<String, String> entry : variables.entrySet()) {
                String placeholder = "x" + entry.getKey() + "x";
                if (text.contains(placeholder)) {
                    text = text.replace(placeholder, removeControlCharacters(entry.getValue()));
                    r.setText(text, 0); // Cập nhật text của run
                }
            }

            // Thay thế placeholder ảnh
            if (text.contains("ximagex") && !imageReplaced) {
                // Xóa placeholder text
                r.setText("", 0); 
                
                File imageFile = new File(qrImagePath);
                if (imageFile.exists()) {
                    try (InputStream is = new FileInputStream(imageFile)) {
                        int format = getImageFormat(qrImagePath);
                        if (format != -1) {
                            r.addPicture(is, format, qrImagePath, IMAGE_WIDTH_EMU, IMAGE_HEIGHT_EMU);
                            imageReplaced = true; // Đặt cờ là đã thay thế ảnh
                        } else {
                            System.err.println("Không thể xác định định dạng ảnh cho: " + qrImagePath);
                        }
                    } catch (IOException e) {
                        System.err.println("Lỗi khi đọc file ảnh " + qrImagePath + ": " + e.getMessage());
                    }
                } else {
                    System.err.println("File ảnh QR không tồn tại: " + qrImagePath);
                }
            }
        }
    }

    public static String removeControlCharacters(String input) {
        if (input == null) {
            return ""; // Trả về chuỗi rỗng thay vì null
        }
        return input.replaceAll("\\p{Cntrl}", "");
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

    // --- Hàm mở file trên desktop (cho tiện kiểm tra) ---
    public static void openDesktop(String filePath) {
        try {
            File file = new File(filePath);
            if (file.exists()) {
                Desktop.getDesktop().open(file);
            } else {
                System.err.println("File not found to open: " + filePath);
            }
        } catch (IOException e) {
            System.err.println("Error opening file: " + e.getMessage());
        }
    }

    // --- Các hàm hỗ trợ sao chép ảnh từ Doc nguồn sang Doc đích (đã có từ ví dụ trước) ---
    private static Map<String, String> copyImagesFromDocToDoc(XWPFDocument sourceDoc, XWPFDocument targetDoc) throws IOException, InvalidFormatException {
        Map<String, String> oldToNewImageIds = new HashMap<>();
        
        List<XWPFPictureData> allPictures = sourceDoc.getAllPictures();
        for (XWPFPictureData picData : allPictures) {
            String oldPicId = sourceDoc.getRelationId(picData);
            if (oldPicId == null) {
                System.err.println("Could not get relation ID for picture: " + picData.getFileName());
                continue;
            }

            // --- Sửa đổi ở đây ---
            // Thay vì picData.getInputStream(), dùng picData.getPackagePart().getInputStream()
            PackagePart packagePart = picData.getPackagePart();
            if (packagePart != null) {
                try (InputStream is = packagePart.getInputStream()) {
                    byte[] bytes = org.apache.poi.util.IOUtils.toByteArray(is); // Sử dụng POI IOUtils
                    int pictureType = picData.getPictureType(); 

                    String newPicId = targetDoc.addPictureData(bytes, pictureType);
                    oldToNewImageIds.put(oldPicId, newPicId);
                } catch (IOException e) {
                    System.err.println("Error reading image data from " + picData.getFileName() + ": " + e.getMessage());
                }
            } else {
                System.err.println("Could not get PackagePart for picture: " + picData.getFileName());
            }
        }
        return oldToNewImageIds;
    }

    private static void updateImageReferencesInParagraph(XWPFParagraph paragraph, Map<String, String> oldToNewImageIds) {
        for (XWPFRun run : paragraph.getRuns()) {
            org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR ctr = run.getCTR();
            for (org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing drawing : ctr.getDrawingList()) {
                for (CTInline inline : drawing.getInlineList()) {
                    CTGraphicalObject graphic = inline.getGraphic();
                    if (graphic != null) {
                        CTGraphicalObjectData graphicData = graphic.getGraphicData();
                        if (graphicData != null) {
                            // --- Sửa đổi ở đây ---
                            // Tạo QName thủ công vì getQName() có thể không được định nghĩa trên SchemaType
                            // Namespace cho CT_Picture là "http://schemas.openxmlformats.org/drawingml/2006/picture"
                            // Tên cục bộ (local name) của phần tử là "pic"
                            QName picQName = new QName("http://schemas.openxmlformats.org/drawingml/2006/picture", "pic");

                            for (org.apache.xmlbeans.XmlObject obj : graphicData.selectChildren(picQName)) {
                                
                                org.openxmlformats.schemas.drawingml.x2006.picture.CTPicture picture = 
                                        (org.openxmlformats.schemas.drawingml.x2006.picture.CTPicture) obj;
                                
                                if (picture != null) {
                                    org.openxmlformats.schemas.drawingml.x2006.main.CTBlip blip = picture.getBlipFill().getBlip();
                                    if (blip != null) {
                                        String oldEmbedId = blip.getEmbed();
                                        if (oldEmbedId != null && oldToNewImageIds.containsKey(oldEmbedId)) {
                                            blip.setEmbed(oldToNewImageIds.get(oldEmbedId));
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    private static void updateImageReferencesInTable(XWPFTable table, Map<String, String> oldToNewImageIds) {
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                for (XWPFParagraph paragraph : cell.getParagraphs()) {
                    updateImageReferencesInParagraph(paragraph, oldToNewImageIds);
                }
            }
        }
    }


    // --- Main method ---
    public static void main(String[] args) {
        String templatePath = HVMain.TEMPLATE_DOCX; // Đảm bảo bạn có file template.docx
        String outputMergedPath = "tonghop.docx";
        String baseQrDirectory = "./BAO CAO/202509/qr"; // Thư mục chứa ảnh QR, ví dụ: "qr/mahv1.png", "qr/mahv2.png"

        
        
        try {
            generateMergedWordFromSqlite(templatePath, outputMergedPath, baseQrDirectory);
            System.out.println("Tạo file Word tổng hợp thành công tại: " + outputMergedPath);
            
            openDesktop(outputMergedPath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    

    
}