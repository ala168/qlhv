import java.awt.Desktop;
import java.io.File; // Cần import này
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;

public class DocxMerger {

     public static void mergeDocxFiles(String inputFolderPath, String outputFolderPath) throws IOException, XmlException, InvalidFormatException {
        Path inputDir = Paths.get(inputFolderPath);
        Path outputDir = Paths.get(outputFolderPath);

        if (!Files.exists(inputDir) || !Files.isDirectory(inputDir)) {
            System.err.println("Input folder does not exist or is not a directory: " + inputFolderPath);
            return;
        }

        // Tạo thư mục đầu ra nếu chưa tồn tại
        if (!Files.exists(outputDir)) {
            Files.createDirectories(outputDir);
        }

        Map<String, List<Path>> filesToMerge = new HashMap<>();

        File inputFolderFile = inputDir.toFile(); // Chuyển Path sang File để duyệt
        File[] listOfFiles = inputFolderFile.listFiles(); // Lấy danh sách các tệp và thư mục

        if (listOfFiles != null) {
            for (File file : listOfFiles) {
                if (file.isFile() && file.getName().toLowerCase().endsWith(".docx")) {
                    Path path = file.toPath(); // Chuyển lại về Path
                    String fileName = path.getFileName().toString();
                    int dotIndex = fileName.indexOf('_');
                    if (dotIndex != -1) {
                        String baseName = fileName.substring(0, dotIndex);
                        List<Path> list = filesToMerge.get(baseName);
                        if (list == null) {
                            list = new ArrayList<>();
                            filesToMerge.put(baseName, list);
                        }
                        list.add(path);
                    }
                }
            }
        } else {
             System.err.println("Input folder is empty or access denied: " + inputFolderPath);
        }


        for (Map.Entry<String, List<Path>> entry : filesToMerge.entrySet()) {
            String baseName = entry.getKey();
            List<Path> docxFiles = entry.getValue();

            if (docxFiles.size() > 1) {
                System.out.println("Merging files for base name: " + baseName);
                XWPFDocument mergedDocument = null; // Khởi tạo null để handle đóng tài liệu
               
                try {
                	String mahv0=HVMain.extractCodeFromPath(docxFiles.get(0).toFile().getAbsolutePath());                             
                    File imagePath0 = new File(outputFolderPath+"/qr",mahv0+".png");
                    Map<String, String> vars0 = new java.util.HashMap<String, String>();
                    String outputPath0 = docxFiles.get(0).toFile().getAbsolutePath();                   
                    WordTemplate.generateWordFromTemplate(outputPath0, outputPath0, vars0, imagePath0.getAbsolutePath());
                    // Lấy tài liệu đầu tiên làm cơ sở
                    try (FileInputStream fis = new FileInputStream(docxFiles.get(0).toFile())) {
                        mergedDocument = new XWPFDocument(fis); 
                    }

                    // Thêm nội dung từ các tài liệu còn lại
                    for (int i = 1; i < docxFiles.size(); i++) {
                        try (FileInputStream fis = new FileInputStream(docxFiles.get(i).toFile())) {
                            //add break
                        	XWPFParagraph pageBreakParagraph = mergedDocument.createParagraph();
                            XWPFRun pageBreakRun = pageBreakParagraph.createRun();
                            pageBreakRun.addBreak(BreakType.PAGE);
                        	
                        	
                        	XWPFDocument nextDoc = new XWPFDocument(fis);

                            CTBody nextDocBody = nextDoc.getDocument().getBody();

                            for (int k = 0; k < nextDocBody.sizeOfPArray(); k++) {
                                mergedDocument.getDocument().getBody().addNewP().set(nextDocBody.getPArray(k));
                            }
                            for (int k = 0; k < nextDocBody.sizeOfTblArray(); k++) {
                                mergedDocument.getDocument().getBody().addNewTbl().set(nextDocBody.getTblArray(k));
                            }
                            
                            
                            Path outputPath = outputDir.resolve(baseName + ".docx");
                            try (FileOutputStream fos = new FileOutputStream(outputPath.toFile())) {
                                if (mergedDocument != null) {
                                     mergedDocument.write(fos);
                                }
                            }
                            
                            String mahv=HVMain.extractCodeFromPath(docxFiles.get(i).toFile().getAbsolutePath());                             
                            File imagePath = new File(outputFolderPath+"/qr",mahv+".png");
                            Map<String, String> vars = new java.util.HashMap<String, String>();
                            WordTemplate.generateWordFromTemplate(outputPath.toString(), outputPath.toString(), vars, imagePath.getAbsolutePath());
                            try (FileInputStream fisMer = new FileInputStream(outputPath.toFile())) {
                                mergedDocument = new XWPFDocument(fisMer); 
                            }
                        }
                    }

                    Path outputPath = outputDir.resolve(baseName + ".docx");
                    try (FileOutputStream fos = new FileOutputStream(outputPath.toFile())) {
                        if (mergedDocument != null) {
                             mergedDocument.write(fos);
                        }
                    }                    

                    //testing
                    if (HVMain.test){
	                    openDesktop(outputPath.toString());
                    }
                } finally {
                    // Đảm bảo đóng mergedDocument nếu nó đã được tạo
                    if (mergedDocument != null) {
                        try {
                            mergedDocument.close();
                        } catch (IOException e) {
                            System.err.println("Error closing merged document for " + baseName + ": " + e.getMessage());
                        }
                    }
                }

            } else if (docxFiles.size() == 1) {
                System.out.println("Only one file for base name '" + baseName + "', skipping merge: " + docxFiles.get(0).getFileName());
                // Tùy chọn: Files.copy(docxFiles.get(0), outputDir.resolve(docxFiles.get(0).getFileName()));
            }
        }
    }
     
     public static void mergeDocxFiles1(String inputFolderPath, String outputFolderPath) throws IOException, XmlException {
         Path inputDir = Paths.get(inputFolderPath);
         Path outputDir = Paths.get(outputFolderPath);

         if (!Files.exists(inputDir) || !Files.isDirectory(inputDir)) {
             System.err.println("Input folder does not exist or is not a directory: " + inputFolderPath);
             return;
         }

         // Tạo thư mục đầu ra nếu chưa tồn tại
         if (!Files.exists(outputDir)) {
             Files.createDirectories(outputDir);
         }

         Map<String, List<Path>> filesToMerge = new HashMap<>();

         File inputFolderFile = inputDir.toFile(); // Chuyển Path sang File để duyệt
         File[] listOfFiles = inputFolderFile.listFiles(); // Lấy danh sách các tệp và thư mục

         if (listOfFiles != null) {
             for (File file : listOfFiles) {
                 if (file.isFile() && file.getName().toLowerCase().endsWith(".docx")) {
                     Path path = file.toPath(); // Chuyển lại về Path
                     String fileName = path.getFileName().toString();
                     int dotIndex = fileName.indexOf('_');
                     if (dotIndex != -1) {
                         String baseName = fileName.substring(0, dotIndex);
                         List<Path> list = filesToMerge.get(baseName);
                         if (list == null) {
                             list = new ArrayList<>();
                             filesToMerge.put(baseName, list);
                         }
                         list.add(path);
                     }
                 }
             }
         } else {
              System.err.println("Input folder is empty or access denied: " + inputFolderPath);
         }


         for (Map.Entry<String, List<Path>> entry : filesToMerge.entrySet()) {
             String baseName = entry.getKey();
             List<Path> docxFiles = entry.getValue();

             if (docxFiles.size() > 1) {
                 System.out.println("Merging files for base name: " + baseName);
                 XWPFDocument mergedDocument = null; // Khởi tạo null để handle đóng tài liệu

                 try {
                     // Lấy tài liệu đầu tiên làm cơ sở
                     try (FileInputStream fis = new FileInputStream(docxFiles.get(0).toFile())) {
                         mergedDocument = new XWPFDocument(fis); 
                     }

                     // Thêm nội dung từ các tài liệu còn lại
                     for (int i = 1; i < docxFiles.size(); i++) {
                         try (FileInputStream fis = new FileInputStream(docxFiles.get(i).toFile())) {
                             //add break
                         	XWPFParagraph pageBreakParagraph = mergedDocument.createParagraph();
                             XWPFRun pageBreakRun = pageBreakParagraph.createRun();
                             pageBreakRun.addBreak(BreakType.PAGE);
                         	
                         	
                         	XWPFDocument nextDoc = new XWPFDocument(fis);

                             CTBody nextDocBody = nextDoc.getDocument().getBody();

                             for (int k = 0; k < nextDocBody.sizeOfPArray(); k++) {
                                 mergedDocument.getDocument().getBody().addNewP().set(nextDocBody.getPArray(k));
                             }
                             for (int k = 0; k < nextDocBody.sizeOfTblArray(); k++) {
                                 mergedDocument.getDocument().getBody().addNewTbl().set(nextDocBody.getTblArray(k));
                             }
                         }
                     }

                     Path outputPath = outputDir.resolve(baseName + ".docx");
                     try (FileOutputStream fos = new FileOutputStream(outputPath.toFile())) {
                         if (mergedDocument != null) {
                              mergedDocument.write(fos);
                         }
                     }
                     //System.out.println("Merged file created: " + outputPath);

                     //testing
                     if (HVMain.test){
 	                    openDesktop(outputPath.toString());
                     }
                 } finally {
                     // Đảm bảo đóng mergedDocument nếu nó đã được tạo
                     if (mergedDocument != null) {
                         try {
                             mergedDocument.close();
                         } catch (IOException e) {
                             System.err.println("Error closing merged document for " + baseName + ": " + e.getMessage());
                         }
                     }
                 }

             } else if (docxFiles.size() == 1) {
                 System.out.println("Only one file for base name '" + baseName + "', skipping merge: " + docxFiles.get(0).getFileName());
                 // Tùy chọn: Files.copy(docxFiles.get(0), outputDir.resolve(docxFiles.get(0).getFileName()));
             }
         }
     }

	public static void openDesktop(String outputPath) {
		Desktop desktop = Desktop.getDesktop();
		try {
		    desktop.open(new File(outputPath.toString()));
		} catch (IOException e) {
		    System.out.println("❌ Lỗi khi mở file: " + e.getMessage());
		}
	}

	public static void openBaseDir(String path) {
		
		File baseDir = new File(path);
		
	    if (baseDir == null || !baseDir.exists() || !baseDir.isDirectory()) {
	        return;
	    }

	    try {
	        if (Desktop.isDesktopSupported()) {
	            Desktop.getDesktop().open(baseDir);
	        } else {
	            String os = System.getProperty("os.name").toLowerCase();
	            Runtime rt = Runtime.getRuntime();

	            if (os.contains("win")) {
	                rt.exec("explorer " + baseDir.getAbsolutePath());
	            } else if (os.contains("mac")) {
	                rt.exec(new String[] { "open", baseDir.getAbsolutePath() });
	            } else if (os.contains("nix") || os.contains("nux")) {
	                rt.exec(new String[] { "xdg-open", baseDir.getAbsolutePath() });
	            } else {
	               
	            }
	        }
	    } catch (IOException e) {
	        e.printStackTrace();	        
	    }
	}
	
    public static void main(String[] args) {
        String inputFolder = ".\\BAO CAO\\202509\\hv";
        String outputFolder = ".\\BAO CAO\\202509";

        try {
        	File baseDir = new File(outputFolder);
        	File[] files = baseDir.listFiles(new FilenameFilter() {
                @Override
                public boolean accept(File dir, String name) {
                    return name.toLowerCase().endsWith(".docx");
                }
            });

            for (File file : files) {
               file.delete();
            }
            
            mergeDocxFiles(inputFolder, outputFolder);
        	openBaseDir(inputFolder);
            System.out.println("Merge process completed.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}