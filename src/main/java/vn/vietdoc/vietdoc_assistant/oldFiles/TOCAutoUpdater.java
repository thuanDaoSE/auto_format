package vn.vietdoc.vietdoc_assistant.oldFiles;

import java.io.BufferedReader;
import java.io.File;
import java.io.InputStreamReader;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.List;

public class TOCAutoUpdater {

    // CẤU HÌNH ĐƯỜNG DẪN LIBREOFFICE
    private static final String LO_PATH_WIN = "C:\\Program Files\\LibreOffice\\program\\soffice.exe";
    private static final String LO_PATH_LINUX = "libreoffice";

    public static void updateDocx(String sourceDocx, String destDocx) {
        System.out.println(">>> [TOC Updater] Bắt đầu quy trình ép cập nhật số trang...");
        
        File source = new File(sourceDocx);
        File dest = new File(destDocx);
        File tempDir = dest.getParentFile();
        if (tempDir != null && !tempDir.exists()) tempDir.mkdirs();

        // BƯỚC 1: Convert DOCX -> ODT (Định dạng OpenOffice)
        // Việc này ép LibreOffice phải "hiểu" và "dàn trang" lại toàn bộ văn bản
        File odtFile = convertFile(source, tempDir, "odt");
        
        if (odtFile != null && odtFile.exists()) {
            System.out.println(">>> [TOC Updater] Bước 1 xong: Đã tạo file trung gian ODT.");
            
            // BƯỚC 2: Convert ODT -> DOCX (Ngược lại)
            // Lúc này số trang đã được tính toán và ghi cứng vào file
            File finalDocx = convertFile(odtFile, tempDir, "docx");
            
            if (finalDocx != null && finalDocx.exists()) {
                try {
                    Files.move(finalDocx.toPath(), dest.toPath(), StandardCopyOption.REPLACE_EXISTING);
                    System.out.println(">>> [TOC Updater] HOÀN TẤT! File chuẩn: " + destDocx);
                    
                    // Dọn dẹp file rác
                    odtFile.delete();
                } catch (Exception e) {
                    e.printStackTrace();
                }
            } else {
                System.err.println("!!! Lỗi: Không thể convert ngược lại DOCX.");
            }
        } else {
            System.err.println("!!! Lỗi: Không thể tạo file trung gian ODT.");
        }
    }

    private static File convertFile(File inputFile, File outputDir, String targetFormat) {
        List<String> command = new ArrayList<>();
        command.add(isWindows() ? LO_PATH_WIN : LO_PATH_LINUX);
        command.add("--headless");
        command.add("--convert-to");
        command.add(targetFormat);
        command.add("--outdir");
        command.add(outputDir.getAbsolutePath());
        command.add(inputFile.getAbsolutePath());

        try {
            ProcessBuilder pb = new ProcessBuilder(command);
            // pb.redirectErrorStream(true); // Uncomment để debug lỗi
            Process process = pb.start();
            
            // Đọc luồng để tránh treo process
            BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream()));
            while (reader.readLine() != null) {} 
            
            int exitCode = process.waitFor();
            if (exitCode == 0) {
                // Tên file kết quả mặc định
                String fileNameNoExt = inputFile.getName().replaceFirst("[.][^.]+$", "");
                return new File(outputDir, fileNameNoExt + "." + targetFormat);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private static boolean isWindows() {
        return System.getProperty("os.name").toLowerCase().contains("win");
    }
}