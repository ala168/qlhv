

	import org.apache.commons.compress.archivers.sevenz.SevenZArchiveEntry;
import org.apache.commons.compress.archivers.sevenz.SevenZOutputFile;

	import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
	
public class Process {

	private static void compressTo7z(String inputFilePath, String outputDirPath) {
	    File inputFile = new File(inputFilePath);
	    if (!inputFile.exists()) {
	        //logMessage("File đầu vào không tồn tại: " + inputFilePath);
	    	System.out.println("File đầu vào không tồn tại: " + inputFilePath);
	        return;
	    }

	    File outputDir = new File(outputDirPath);
	    if (!outputDir.exists()) {
	        outputDir.mkdirs();
	    }

	    // tạo tên file output với timestamp
	    String baseName = inputFile.getName();
	    int dotIndex = baseName.lastIndexOf(".");
	    if (dotIndex > 0) {
	        baseName = baseName.substring(0, dotIndex);
	    }
	    String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
	    File outputFile = new File(outputDir, baseName + "_" + timestamp + ".7z");

	    SevenZOutputFile sevenZOutput = null;
	    FileInputStream fis = null;
	    try {
	        sevenZOutput = new SevenZOutputFile(outputFile);

	        SevenZArchiveEntry entry = sevenZOutput.createArchiveEntry(inputFile, inputFile.getName());
	        sevenZOutput.putArchiveEntry(entry);

	        fis = new FileInputStream(inputFile);
	        byte[] buffer = new byte[8192];
	        int len;
	        while ((len = fis.read(buffer)) != -1) {
	            sevenZOutput.write(buffer, 0, len);
	        }

	        sevenZOutput.closeArchiveEntry();
	        System.out.println("Đã nén thành công: " + outputFile.getAbsolutePath());

	    } catch (IOException e) {
	        e.printStackTrace();
	        //logMessage("Lỗi khi nén file: " + e.getMessage());
	    } finally {
	        try { if (fis != null) fis.close(); } catch (Exception ignored) {}
	        try { if (sevenZOutput != null) sevenZOutput.close(); } catch (Exception ignored) {}
	    }
	}
	
	public static void main(String[] args) {
		compressTo7z("./QLHV/data/hvtest.db", "./backup");
	}

}