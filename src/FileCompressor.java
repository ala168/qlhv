import java.io.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class FileCompressor {

    /**
     * Nén 1 file thành file zip
     * @param inputPath  đường dẫn file gốc
     * @param outputPath đường dẫn file zip đầu ra
     */
    public static boolean compressFile(String inputPath, String outputPath) {
        File inputFile = new File(inputPath);
        if (!inputFile.exists() || !inputFile.isFile()) {
            System.out.println("File không tồn tại: " + inputPath);
            return false;
        }

        byte[] buffer = new byte[4096]; // buffer 4KB

        FileOutputStream fos = null;
        ZipOutputStream zos = null;
        FileInputStream fis = null;

        try {
            fos = new FileOutputStream(outputPath);
            zos = new ZipOutputStream(fos);

            fis = new FileInputStream(inputFile);
            ZipEntry zipEntry = new ZipEntry(inputFile.getName());
            zos.putNextEntry(zipEntry);

            int len;
            while ((len = fis.read(buffer)) > 0) {
                zos.write(buffer, 0, len);
            }

            zos.closeEntry();
            System.out.println("Đã backup file: " + inputPath + " -> " + outputPath);
            return true;
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        } finally {
            try { if (fis != null) fis.close(); } catch (Exception ignored) {}
            try { if (zos != null) zos.close(); } catch (Exception ignored) {}
            try { if (fos != null) fos.close(); } catch (Exception ignored) {}
        }
    }

    // Test
    public static void main(String[] args) {
        compressFile("E:/workspace/profile_workspace/QLHV/QLHV/data/hvtest.db", "E:/workspace/profile_workspace/QLHV/QLHV/data/hvtest.zip");
    }
}
