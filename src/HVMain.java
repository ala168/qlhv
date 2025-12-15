import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;

import javax.swing.*;

import java.awt.*;
import java.awt.Font;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.*;
import java.text.DecimalFormat;
import java.text.Normalizer;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
//qr
import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLEncoder;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Map;

@SuppressWarnings("serial")
public class HVMain extends JFrame {
    /************************************************/
    public static boolean hasProxy = false;
    public static boolean test = false;
    /************************************************/
    
	public static final String ROOT_DIR = ".";
    public static final String TEMPLATE_DOCX = ROOT_DIR+"/data/template.docx";
    public static final String DB_PATH = ROOT_DIR+"/data/hvtest.db";
    public static final String DB_BK_PATH = ROOT_DIR+"/data/backup";
    public static final String URL_DB = "jdbc:sqlite:"+DB_PATH;
    
    public static final String TEMPLATE_EXCEL_DETAL = ROOT_DIR+"/data/File chi iet tung lop.xlsx";
    public static final String TEMPLATE_EXCEL_BANK = "";
    
	public JTextField folderPathField;
    public JButton chooseButton, btnImportDataButton;
    public JProgressBar progressBar;
    public JTextArea logArea;
    public JSpinner monthSpinner;
    public JButton btnExportWordButton, btnExportExcelButton;
    public JButton btnImportNHButton;
	public JButton btnOpenFolderButton;
	public JCheckBox proxyCheckBox;
	public JButton btnExportGVButton;
	public JButton btnImportSKButton;
    
    public HVMain() {
        setTitle("QLHV");
        setSize(1296, 864);
        Font font = new Font("Tahoma", Font.BOLD, 14);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLayout(new BorderLayout(10, 10)); // Border layout cho JFrame chính

        // 1. Panel trên cùng (NORTH) - Giữ nguyên hoặc điều chỉnh nhẹ
        JPanel topPanel = new JPanel(new GridLayout(2, 1, 5, 5)); // GridLayout cho 2 hàng
        topPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 5, 10)); // Thêm padding

        // Hàng 1: Chọn tháng
        JPanel monthSelectionPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 12, 0)); // FlowLayout trái
        SpinnerDateModel model = new SpinnerDateModel(new Date(), null, null, java.util.Calendar.MONTH);
        monthSpinner = new JSpinner(model);
        JSpinner.DateEditor editor = new JSpinner.DateEditor(monthSpinner, "yyyyMM");
        monthSpinner.setEditor(editor);
        monthSpinner.setFont(font);
        monthSpinner.setPreferredSize(new Dimension(120, 28)); // 120px ngang, 28px cao
        monthSelectionPanel.add(new JLabel("Nhập tháng yyyyMM"));
        monthSelectionPanel.setFont(font);
        monthSelectionPanel.add(monthSpinner);
        
        btnOpenFolderButton = new JButton("1. Mở thư mục");
        btnOpenFolderButton.setPreferredSize(new Dimension(150, 30));
        btnOpenFolderButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent e) {
                new Thread(new Runnable() {
                    public void run() {
                    	SimpleDateFormat sdf = new SimpleDateFormat("yyyyMM");
                        int thangInt = Integer.parseInt(sdf.format((Date) monthSpinner.getValue()));

                        File baocaoDir = new File("BAO CAO");
                        if (!baocaoDir.exists()) baocaoDir.mkdirs();

                        File baseDir = new File(baocaoDir, String.valueOf(thangInt));
                        if (!baseDir.exists()) baseDir.mkdirs();

                        File excelDir = new File(baseDir, "excels");
                        if (!excelDir.exists()) excelDir.mkdirs();
                        
                        
                        
                    	DocxMerger.openDesktop(baseDir.getAbsolutePath());
                    }
                }).start();
            }
        });
        monthSelectionPanel.add(btnOpenFolderButton);
        
        // Tạo CheckBox
        proxyCheckBox = new JCheckBox("Dùng Proxy");
        proxyCheckBox.setVisible(hasProxy);
        proxyCheckBox.setSelected(hasProxy);
        proxyCheckBox.addActionListener(new java.awt.event.ActionListener() {
            @Override
            public void actionPerformed(java.awt.event.ActionEvent e) {
                hasProxy = proxyCheckBox.isSelected();
                init();
            }
        });
        monthSelectionPanel.add(proxyCheckBox, BorderLayout.EAST);
        
        JLabel excelLink = new JLabel("<html><a href=''>Mẫu Excel chi tiết</a></html>");
        excelLink.setCursor(new Cursor(Cursor.HAND_CURSOR));
        //excelLink.setForeground(new Color());
        excelLink.addMouseListener(new java.awt.event.MouseAdapter() {
            @Override
            public void mouseClicked(java.awt.event.MouseEvent e) {
               try {
					Desktop.getDesktop().open(new File(HVMain.TEMPLATE_EXCEL_DETAL));
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
            }
        });  
        monthSelectionPanel.add(excelLink, BorderLayout.EAST);
        
        topPanel.add(monthSelectionPanel);
        
        // Hàng 2: Chọn thư mục
        JPanel folderSelectionPanel = new JPanel(new BorderLayout(5, 5));
        folderPathField = new JTextField();
        folderPathField.setEditable(false);
        chooseButton = new JButton("2.Chọn thư mục");
        chooseButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent e) {
            	setUIComponentsEnabled(false);
            	progressBar.setIndeterminate(false);               
                btn2ChooseFolder_Event();
                setUIComponentsEnabled(true);
                progressBar.setEnabled(true);
            }
        });
        folderSelectionPanel.add(new JLabel("2.Chọn Thư mục các files "), BorderLayout.WEST);
        folderSelectionPanel.add(folderPathField, BorderLayout.CENTER);
        folderSelectionPanel.add(chooseButton, BorderLayout.EAST);
        topPanel.add(folderSelectionPanel);

        add(topPanel, BorderLayout.NORTH);

        // 2. Log Area ở giữa (CENTER)
        logArea = new JTextArea();
        logArea.setEditable(false);
       
        logArea.setFont(font);
        JScrollPane scrollPane = new JScrollPane(logArea);
        scrollPane.setBorder(BorderFactory.createEmptyBorder(0, 10, 0, 10)); // Thêm padding ngang
        add(scrollPane, BorderLayout.CENTER);

        // 3. Panel dưới cùng (SOUTH) - Chứa các button và Progress Bar
        JPanel bottomContainerPanel = new JPanel(new BorderLayout(10, 10));
        bottomContainerPanel.setBorder(BorderFactory.createEmptyBorder(5, 10, 10, 10)); // Thêm padding

        // Panel cho các button
        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.CENTER, 15, 5)); // Các nút ở giữa, cách nhau 15px

        btnImportDataButton = new JButton("3.Đọc file excels");
        btnImportDataButton.setPreferredSize(new Dimension(180, 40));
        btnImportDataButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent e) {
                new Thread(new Runnable() {
                    public void run() {                  	
                    	setUIComponentsEnabled(false);
                        sendLogInfo(btnImportDataButton.getText());  
                        btn3ImportData_Event();
                        setUIComponentsEnabled(true);
                    }
                }).start();
            }
        });
        buttonPanel.add(btnImportDataButton);

        btnImportNHButton = new JButton("4.Nhập TK NH");
        btnImportNHButton.setPreferredSize(new Dimension(180, 40));
        btnImportNHButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent e) {
                new Thread(new Runnable() {
                    public void run() {
                    	setUIComponentsEnabled(false);
                        sendLogInfo(btnImportNHButton.getText());
                        btn4ImportNganHang_Event();
                        setUIComponentsEnabled(true);
                    }
                }).start();
            }
        });
        buttonPanel.add(btnImportNHButton);

        
        btnExportExcelButton = new JButton("5.Tạo Excel tổng");
        btnExportExcelButton.setPreferredSize(new Dimension(180, 40));
        btnExportExcelButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent e) {
                new Thread(new Runnable() {
                    public void run() {
                    	setUIComponentsEnabled(false);
                    	sendLogInfo(btnExportExcelButton.getText());
                        String file=btn5ExportToSumaryExcel_Event(); 
                        if (!file.isEmpty()) DocxMerger.openDesktop(file);
                        setUIComponentsEnabled(true);
                    }
                }).start();
            }
        });
        buttonPanel.add(btnExportExcelButton);
        
        btnExportWordButton = new JButton("6.Xuất Word");
        btnExportWordButton.setPreferredSize(new Dimension(180, 40));
        btnExportWordButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent e) {
                new Thread(new Runnable() {
                    public void run() {
                    	setUIComponentsEnabled(false);
                        sendLogInfo(btnExportWordButton.getText());
                        btn6ExportWordFiles_Event(); // Giả định phương thức này tồn tại
                        setUIComponentsEnabled(true);
                    }
                }).start();
            }
        });
        buttonPanel.add(btnExportWordButton);
        
        
        btnExportGVButton = new JButton("7.Tổng hợp GV");
        btnExportGVButton.setPreferredSize(new Dimension(180, 40));
        btnExportGVButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent e) {
                new Thread(new Runnable() {
                    public void run() {
                    	setUIComponentsEnabled(false);
                        sendLogInfo(btnExportGVButton.getText());
                        btn7TongHopGV_Event(); // Giả định phương thức này tồn tại
                        setUIComponentsEnabled(true);
                    }
                }).start();
            }
        });
        buttonPanel.add(btnExportGVButton);
        
        
        btnImportSKButton = new JButton("8.Nhập sao kê");
        btnImportSKButton.setPreferredSize(new Dimension(180, 40));
        btnImportSKButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent e) {
                new Thread(new Runnable() {
                    public void run() {
                    	setUIComponentsEnabled(false);
                        sendLogInfo(btnImportSKButton.getText());
                    	btn8ImportSaoke_Event();
                        setUIComponentsEnabled(true);
                    }
                }).start();
            }
        });
        buttonPanel.add(btnImportSKButton);
        

        bottomContainerPanel.add(buttonPanel, BorderLayout.CENTER); // Button panel ở giữa của bottomContainerPanel

        // Progress bar dưới các button
        progressBar = new JProgressBar();
        progressBar.setMinimum(0);
        progressBar.setString("");
        progressBar.setStringPainted(true);
        bottomContainerPanel.add(progressBar, BorderLayout.SOUTH); // Progress bar ở dưới của bottomContainerPanel
        
        add(bottomContainerPanel, BorderLayout.SOUTH); // Đặt bottomContainerPanel vào SOUTH của JFrame
        
        
        // 1. Tạo JLabel cho footer với căn lề phải
        JLabel footerLabel = new JLabel("v1.104 Designed by tiepvk   ");
        footerLabel.setHorizontalAlignment(SwingConstants.RIGHT);

        // 2. Tạo một panel chính mới cho toàn bộ khu vực SOUTH
        JPanel mainSouthPanel = new JPanel(new BorderLayout());

        // 3. Thêm panel chứa button và progress bar vào giữa panel chính này
        mainSouthPanel.add(bottomContainerPanel, BorderLayout.CENTER);

        // 4. Thêm footer label vào dưới cùng của panel chính
        mainSouthPanel.add(footerLabel, BorderLayout.SOUTH);

        // 5. Thêm panel chính (đã chứa cả button và footer) vào SOUTH của JFrame
        add(mainSouthPanel, BorderLayout.SOUTH);
        
        
        setComponentFontRecursively(this.getContentPane(), font); 
        
        footerLabel.setFont(new Font("Tahoma",Font.ITALIC, 10));
        proxyCheckBox.setFont(new Font("Tahoma",Font.PLAIN, 12));
        
        setLocationRelativeTo(null); // Đặt cửa sổ ra giữa màn hình
        setVisible(true); // Hiển thị JFrame
        
        init();
    }

    
    public void setComponentFontRecursively(Container container, Font font) {
        for (Component comp : container.getComponents()) {
            if (comp instanceof JLabel || comp instanceof JButton) {
                comp.setFont(font);
            }
            if (comp instanceof Container) {
                setComponentFontRecursively((Container) comp, font);
            }
        }
    }
    
    public void init(){
    	if (hasProxy) {
    		System.setProperty("https.proxyHost", "10.53.120.1");
            System.setProperty("https.proxyPort", "8080");
            
            System.setProperty("http.proxyHost", "10.53.120.1");
            System.setProperty("http.proxyPort", "8080");
        			
            System.setProperty("socksProxyHost", "10.53.120.1");
            System.setProperty("socksProxyPort", "8080");
        }
    	
        HtmlReader.readHtmlFromSite();
        
        if (!HtmlReader.readOK){
        	HtmlReader.enable = false;
        	setUIComponentsEnabled(HtmlReader.enable);
        	logMessage("Kiểm tra lại kết nối Internet và Mở lại ứng dụng!");
			JOptionPane.showMessageDialog(this, "Kiểm tra lại kết nối Internet và Mở lại ứng dụng!");
			proxyCheckBox.setVisible(true);
        	return;
        }        
        
        if (HtmlReader.hasUpdate) {
        	try {
				if (HtmlReader.update()) {
					HtmlReader.enable = false;
					setUIComponentsEnabled(HtmlReader.enable);
					logMessage("Đã update thành công, hãy tắt ứng dụng và mở lại!");
					JOptionPane.showMessageDialog(this, "Đã update thành công, hãy tắt ứng dụng và mở lại!");
					return;
				}
			} catch (Exception e) {
				
			}
        }
        
        setUIComponentsEnabled(HtmlReader.enable);
        if (HtmlReader.enable) {
        	logMessage("Chào mừng đến với ứng dụng QLHV!", true);
        	logMessage("Để bắt đầu hãy:");
        	logMessage("1. Chọn Ngày tháng => nhấn "+ btnOpenFolderButton.getText() );
        	logMessage("2. Copy các file chi tiết vào thư mục excels => nhấn nút "+ chooseButton.getText());  
        	
        	SimpleDateFormat sdf = new SimpleDateFormat("yyyyMM");
        	int thangInt = Integer.parseInt(sdf.format((Date) monthSpinner.getValue()));
        	if (checkTranDataExists(thangInt, true)) {
        		logMessage("-------------------------------------------");
        		logMessage("Dữ liệu tổng hợp tháng " + thangInt + " đang được thanh toán!");
        		logMessage("=> Có thể chọn " + btnExportExcelButton.getText() + " ...->... " + btnImportSKButton.getText());
        	} else {
        		if (checkTranDataExists(thangInt, false)) {
            		logMessage("-------------------------------------------");
            		logMessage("Đang có dữ liệu tổng hợp và chưa có thông tin thanh toán tháng " + thangInt);
            		logMessage("=> Có thể chọn " + btnExportExcelButton.getText() + " ...->... " + btnImportSKButton.getText());
            	}
        	}
        }
        
        //fixOnlyDuplicateMahv();
        makeMahvUniqueSimple();
    }
    
    public void makeMahvUniqueSimple() {
        Connection conn = null;
        Statement st = null;

        try {
            conn = DriverManager.getConnection(URL_DB);
            st = conn.createStatement();

            st.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_hv_mahv ON hv(mahv)");
            st.execute("CREATE INDEX IF NOT EXISTS idx_hv_loptt ON hv(loptt)");

        } catch (Exception e) {
            e.printStackTrace();
            logMessage("Lỗi makeMahvUniqueSimple: " + e.getMessage());
            sendLogException(e);
        } finally {
            try { if (st != null) st.close(); } catch (Exception e) {}
            try { if (conn != null) conn.close(); } catch (Exception e) {}
        }
    }
    
    public void logMessage(final String message) {
        SwingUtilities.invokeLater(new Runnable() {
            public void run() {
                logArea.append(message + "\n");
                logArea.setCaretPosition(logArea.getDocument().getLength());
            }
        });
    }
    
    public void logMessage(final String message, final boolean reset) {
        SwingUtilities.invokeLater(new Runnable() {
            public void run() {
            	if (reset) logArea.setText("");
                logArea.append(message + "\n");
                logArea.setCaretPosition(logArea.getDocument().getLength());
            }
        });
    }

    public void btn8ImportSaoke_Event() {
    	// Lấy tháng từ monthSpinner để tạo baseDir
    	SimpleDateFormat sdf = new SimpleDateFormat("yyyyMM");
    	int thangInt = Integer.parseInt(sdf.format((Date) monthSpinner.getValue()));
    	
    	//kiem tra chua co du lieu tong hop thang thi quit
    	if (!checkTranDataExists(thangInt, false)) {
    		logMessage("Chưa có dữ liệu tổng hợp tháng!");
    		logMessage("=> Hãy chọn " + btnImportDataButton.getText() + " hoặc " + btnImportNHButton.getText());
    		JOptionPane.showMessageDialog(this, "Chưa có dữ liệu tổng hợp tháng!");
    		return;
    	}
    	
    	File baocaoDir = new File("BAO CAO");
    	File baseDir = new File(baocaoDir, String.valueOf(thangInt));

    	// JFileChooser mặc định tới baseDir
    	JFileChooser chooser = new JFileChooser();
    	chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);

    	if (baseDir.exists()) {
    	    chooser.setCurrentDirectory(baseDir);
    	} else {
    	    chooser.setCurrentDirectory(baocaoDir.exists() ? baocaoDir : new File("."));
    	}

    	int result = chooser.showOpenDialog(this);
    	if (result != JFileChooser.APPROVE_OPTION) {
    	    return;
    	}

    	File excelFilePath = chooser.getSelectedFile();
    	if (!excelFilePath.getName().toLowerCase().endsWith(".xlsx")) {
    	    JOptionPane.showMessageDialog(this, "Chỉ chọn file Excel .xlsx");
    	    return;
    	}

    	// Nếu cần: lưu đường dẫn đã chọn vào text field
    	folderPathField.setText(excelFilePath.getAbsolutePath());
    	logMessage("Bắt đầu nhập dữ liệu sao kê ngân hàng ...", true);

    	boolean readSKOK = readSaoKe(excelFilePath);
    	
    	if (!readSKOK) return;
    	
    	addMissingColumnsToTongHopThang();
    	
    	updateTongHopThangFromSaoKe();
    	
    	String file=btn5ExportToSumaryExcel_Event(); 
    	exportSaoKeMissingMaHV(file);
        if (!file.isEmpty()) DocxMerger.openDesktop(file);
    }
    
    public void exportSaoKeMissingMaHV(String file) {
        Connection conn = null;
        PreparedStatement ps = null;
        ResultSet rs = null;

        try {
            // Lấy tháng từ monthSpinner
            SimpleDateFormat sdf = new SimpleDateFormat("yyyyMM");
            int thang = Integer.parseInt(sdf.format((Date) monthSpinner.getValue()));

			File rootDir = new File("BAO CAO");
			if (!rootDir.exists()) rootDir.mkdirs();
			
			File baseDir = new File(rootDir, String.valueOf(thang));
			if (!baseDir.exists()) baseDir.mkdirs();

			String outputPath = file;
			if (outputPath==null || outputPath.isEmpty()) {
				outputPath = new File(baseDir, "SaoKe_KhongCoMaHV_" + thang + ".xlsx").getAbsolutePath();
			}

            conn = DriverManager.getConnection(URL_DB);

            // Truy vấn lấy dữ liệu mahv trống
            String sql = "SELECT sotk, trandate, transactionNo, tranamt, remark, mahv " +
                         "FROM SaoKeThang " +
                         "WHERE mahv IS NULL OR TRIM(mahv)='' or mahv not in (select upper(mahv) from tonghopthang)";

            ps = conn.prepareStatement(sql);
            rs = ps.executeQuery();

            // Tạo workbook
            FileInputStream fis = new FileInputStream(outputPath);            
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            
            fis.close(); // nên đóng ngay sau khi load
            XSSFSheet sheet = wb.createSheet("SaoKeThieuMaHV");

            // Tạo dòng header
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Ngay GD");
            header.createCell(1).setCellValue("So GD");
            header.createCell(2).setCellValue("So Tien");
            header.createCell(3).setCellValue("Noi dung");
            header.createCell(4).setCellValue("So TK");
            header.createCell(5).setCellValue("Mã HS");

            // Ghi dữ liệu
            int rowIndex = 1;
            while (rs.next()) {
                Row row = sheet.createRow(rowIndex++);

                row.createCell(0).setCellValue(rs.getString("trandate"));
                row.createCell(1).setCellValue(rs.getString("transactionNo"));
                row.createCell(2).setCellValue(rs.getDouble("tranamt"));
                row.createCell(3).setCellValue(rs.getString("remark"));
                row.createCell(4).setCellValue(rs.getString("sotk"));
                row.createCell(5).setCellValue(rs.getString("mahv"));
            }

            // Auto-size cột
            for (int i = 0; i < 6; i++) {
                sheet.autoSizeColumn(i);
            }

            // Ghi file Excel
            FileOutputStream fos = new FileOutputStream(outputPath);
            wb.write(fos);
            fos.close();
            wb.close();
            
            logMessage("=>File dữ liệu KHÔNG xác định được mã HS: " + outputPath);

        } catch (Exception e) {
            e.printStackTrace();
            logMessage("Lỗi khi tạo file: " + e.toString());
            sendLogException(e);
        } finally {
            try { if (rs != null) rs.close(); } catch (Exception e) {}
            try { if (ps != null) ps.close(); } catch (Exception e) {}
            try { if (conn != null) conn.close(); } catch (Exception e) {}
        }
    }

    public boolean readSaoKe(File excelFilePath) {
        Connection conn = null;
        PreparedStatement psCheck = null;
        PreparedStatement psInsert = null;
        PreparedStatement psSelectRemark = null;
        PreparedStatement psUpdateRemark = null;
        Statement st = null;
        ResultSet rs = null;

        try {
            conn = DriverManager.getConnection(URL_DB);
            st = conn.createStatement();

            // Tạo bảng nếu chưa tồn tại
            st.execute("CREATE TABLE IF NOT EXISTS SaoKeThang (" +
                    "id INTEGER PRIMARY KEY AUTOINCREMENT," +
                    "sotk TEXT," +
                    "trandate TEXT," +
                    "transactionNo TEXT," +
                    "tranamt REAL," +
                    "Remark TEXT," +
                    "month INTEGER," +
                    "mahv TEXT," +
                    "tenhv TEXT" +
                    ")");

            addMissingColumnsToSaoKeThang();
            
            st.execute("CREATE INDEX IF NOT EXISTS idx_saoke_transactionno ON SaoKeThang(transactionNo, sotk)");
            //st.execute("Delete from SaoKeThang");

            // kiểm tra trùng
            psCheck = conn.prepareStatement(
                    "SELECT COUNT(*) FROM SaoKeThang WHERE transactionNo = ? AND sotk = ?"
            );

            // Lấy Remark cũ
            psSelectRemark = conn.prepareStatement(
                    "SELECT Remark FROM SaoKeThang WHERE transactionNo = ? AND sotk = ?"
            );

            // Update Remark nếu khác
            psUpdateRemark = conn.prepareStatement(
                    "UPDATE SaoKeThang SET Remark=?,mahv=? WHERE transactionNo=? AND sotk=?"
            );

            // insert
            psInsert = conn.prepareStatement(
                    "INSERT INTO SaoKeThang (sotk, month, trandate, transactionNo, tranamt, Remark, mahv, tenhv, phitm) " +
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)"
            );

            // Đọc Excel
            FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0);

            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

            boolean isSaoKeCaNhan = true;

            Row row4 = sheet.getRow(3);
            Cell cellF1 = row4.getCell(16);
            String tmpAccountNo = getCellValue(cellF1, evaluator).trim();

            if (!tmpAccountNo.isEmpty() && tmpAccountNo.toUpperCase().contains("ACCOUNT NO"))
                isSaoKeCaNhan = false;
            else {
                Row row8 = sheet.getRow(7);
                tmpAccountNo = getCellValue(row8.getCell(1), evaluator).trim();
            }

            String soTk = StringUtils.extractAccountNumber(tmpAccountNo);

            int monthValue = 0;
            int startRow = 1;
            int lastRow = sheet.getLastRowNum();
            int insertedCount = 0;
            int insertedNoMaCount = 0;
            boolean isStartRow = false;

            for (int i = startRow; i <= lastRow; i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

				//getCellValue(row.getCell(0),evaluator).trim(); // cột A
                //getCellValue(row.getCell(1),evaluator).trim(); // cột B
                //getCellValue(row.getCell(2),evaluator).trim(); // cột C
                //getCellValue(row.getCell(3),evaluator).trim(); // cột D
                //getCellValue(row.getCell(4),evaluator).trim(); // cột E
                //getCellValue(row.getCell(5),evaluator).trim(); // cột F
				//getCellValue(row.getCell(6),evaluator).trim(); // cột G
				//getCellValue(row.getCell(7),evaluator).trim(); // cột H
                //getCellValue(row.getCell(9),evaluator).trim(); // cột J
                //getCellValue(row.getCell(10),evaluator).trim(); // cột K
                //getCellValue(row.getCell(11),evaluator).trim(); // cột L
				//getCellValue(row.getCell(13),evaluator).trim(); // cột N
                //getCellValue(row.getCell(14),evaluator).trim(); // cột O
                //getCellValue(row.getCell(15),evaluator).trim(); // cột P
                //getCellValue(row.getCell(16),evaluator).trim(); // cột Q

                int col = isSaoKeCaNhan ? 1 : 0;
                String STT = getCellValue(row.getCell(col), evaluator).trim();

                if (STT.isEmpty()) continue;
                if (STT.toUpperCase().contains("STT")) isStartRow = true;
                if (!isStartRow) continue;

                col = isSaoKeCaNhan ? 4 : 6;
                String date = getCellValue(row.getCell(col), evaluator).trim();

                col = isSaoKeCaNhan ? 10 : 8;
                String creditStr = getCellValue(row.getCell(col), evaluator).trim();

                col = isSaoKeCaNhan ? 11 : 12;
                String remark = getCellValue(row.getCell(col), evaluator).trim();

                col = isSaoKeCaNhan ? 6 : 17;
                String transactionNo = getCellValue(row.getCell(col), evaluator).trim();

                col = isSaoKeCaNhan ? 9 : 7;
                String debitStr = getCellValue(row.getCell(col), evaluator).trim();
                
                if (date.isEmpty() || transactionNo.isEmpty() || creditStr.isEmpty() || remark.isEmpty())
                    continue;

                String mahv = "";
                String tenhv = "";
                String thang = "0";

                StringUtils.RemarkInfo info = StringUtils.parseRemark(remark);
                if (info != null) {
                    mahv = info.mahv;
                    thang = String.valueOf(info.month);
                    try { monthValue = Integer.valueOf(thang); } catch (Exception e1) {}
                    tenhv = info.tenhv;
                } else {
                    monthValue = 0;
                }

                // ------------ KIỂM TRA TỒN TẠI ------------
                psCheck.setString(1, transactionNo);
                psCheck.setString(2, soTk);
                rs = psCheck.executeQuery();
                boolean exists = rs.next() && rs.getInt(1) > 0;
                rs.close();

                if (exists) {
                	
                	if (mahv.isEmpty()) {
                		insertedNoMaCount++;
                		logMessage("Đã nhập lần trước số tk " + soTk + " mã = " + transactionNo + ", ngày=" + date);
                		continue;
                	}
                	
                    // LẤY REMARK CŨ ĐỂ SO SÁNH
                    psSelectRemark.setString(1, transactionNo);
                    psSelectRemark.setString(2, soTk);
                    rs = psSelectRemark.executeQuery();

                    if (rs.next()) {
                        String oldRemark = rs.getString("Remark");
                        if (oldRemark == null) oldRemark = "";

                        if (!oldRemark.equals(remark)) {
                            // UPDATE remark mới
                            psUpdateRemark.setString(1, remark);
                            psUpdateRemark.setString(2, mahv);
                            psUpdateRemark.setString(3, transactionNo);
                            psUpdateRemark.setString(4, soTk);
                            psUpdateRemark.executeUpdate();
                            insertedCount++;
                            
                            logMessage("Cập nhật cho giao dịch " + transactionNo + " (" + soTk + ")");
                        } else {
                        	logMessage("Đã nhập lần trước số tk " + soTk + " mã = " + transactionNo + ", ngày=" + date);
                            
                        }
                    }
                    rs.close();
                    
                    continue;
                }

                // ---------------- INSERT MỚI ----------------
                double tranAmt = parseDoubleSafe(creditStr);
                if (tranAmt == 0) continue;
                double phitm = parseDoubleSafe(debitStr);

                psInsert.setString(1, soTk);
                psInsert.setInt(2, monthValue);
                psInsert.setString(3, date);
                psInsert.setString(4, transactionNo);
                psInsert.setDouble(5, tranAmt);
                psInsert.setString(6, remark);
                psInsert.setString(7, mahv);
                psInsert.setString(8, tenhv);
                psInsert.setDouble(9, phitm);

                psInsert.executeUpdate();
                insertedCount++;
                if (mahv.isEmpty()) insertedNoMaCount++;
            }

            workbook.close();
            fis.close();

            logMessage("Đã hoàn tất đọc sao kê. Đã THÊM/Cập nhập " + insertedCount +
                    " dòng, trong đó " + insertedNoMaCount + " KHÔNG xác định mã HV.");

            return true;

        } catch (Exception e) {
            e.printStackTrace();
            logMessage("Lỗi importBankStatement: " + e.getMessage());
            sendLogException(e);
            return false;
        } finally {
            try { if (rs != null) rs.close(); } catch (Exception e) {}
            try { if (psCheck != null) psCheck.close(); } catch (Exception e) {}
            try { if (psInsert != null) psInsert.close(); } catch (Exception e) {}
            try { if (psSelectRemark != null) psSelectRemark.close(); } catch (Exception e) {}
            try { if (psUpdateRemark != null) psUpdateRemark.close(); } catch (Exception e) {}
            try { if (st != null) st.close(); } catch (Exception e) {}
            try { if (conn != null) conn.close(); } catch (Exception e) {}
        }
    }

    
	public boolean readSaoKe2_bkKhongDung(File excelFilePath) {
		Connection conn = null;
        PreparedStatement psCheck = null;
        PreparedStatement psInsert = null;
        Statement st = null;
        ResultSet rs = null;

        try {
            conn = DriverManager.getConnection(URL_DB);
            st = conn.createStatement();
            
            // Tạo bảng nếu chưa tồn tại
            st.execute("CREATE TABLE IF NOT EXISTS SaoKeThang (" +
                    "id INTEGER PRIMARY KEY AUTOINCREMENT," +
                    "sotk TEXT," +                    
                    "trandate TEXT," +
                    "transactionNo TEXT," +
                    "tranamt REAL," +                    
                    "Remark TEXT,"+
                    "month INTEGER," +
                    "mahv TEXT,"+
                    "tenhv TEXT" +
                    ")");

            st.execute("CREATE INDEX IF NOT EXISTS idx_saoke_transactionno ON SaoKeThang(transactionNo,sotk)");
            //st.execute("Delete from SaoKeThang");

            // Chuẩn bị câu lệnh kiểm tra trùng
            psCheck = conn.prepareStatement(
                    "SELECT COUNT(*) FROM SaoKeThang WHERE transactionNo = ? AND sotk = ?"
            );

            // Chuẩn bị câu lệnh insert
            psInsert = conn.prepareStatement(
                    "INSERT INTO SaoKeThang (sotk, month, trandate, transactionNo, tranamt, Remark, mahv, tenhv) " +
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?)"
            );

            // Đọc Excel
            FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0);

            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

            
            boolean isSaoKeCaNhan = true;
            // Lấy số tài khoản từ ô F1
            Row row4 = sheet.getRow(3);
            Cell cellF1 = row4.getCell(16); //Q4
            String tmpAccountNo = getCellValue(cellF1,evaluator).trim();
            if (!tmpAccountNo.isEmpty() && tmpAccountNo.toUpperCase().contains("Account No".toUpperCase())) isSaoKeCaNhan = false;
            else {
            	Row row8 = sheet.getRow(7);//B2
            	tmpAccountNo = getCellValue(row8.getCell(1),evaluator).trim();            	
            }
            
            if (test) System.out.println(tmpAccountNo);
            String soTk = StringUtils.extractAccountNumber(tmpAccountNo);
            if (test) System.out.println("soTk="+soTk);

            int monthValue = 0;
            int startRow = 1; // Excel row 5 → index = 4
            int lastRow = sheet.getLastRowNum();
            int insertedCount = 0;
            int insertedNoMaCount = 0;
            boolean isStartRow = false;
            for (int i = startRow; i <= lastRow; i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;
                
                //STT
                int tmpi = 0;
                if (isSaoKeCaNhan) tmpi = 1;
                String STT = getCellValue(row.getCell(tmpi),evaluator).trim();  // cột A
                
                if (test) System.out.println(i + ":" + STT);
                
                if (STT.isEmpty()) continue;
                if (STT.toUpperCase().contains("STT")) isStartRow = true;
                if (!isStartRow) continue;

                //getCellValue(row.getCell(0),evaluator).trim();  // cột A
                //getCellValue(row.getCell(1),evaluator).trim(); // cột B
                //getCellValue(row.getCell(2),evaluator).trim(); // cột C
                //getCellValue(row.getCell(3),evaluator).trim(); // cột D
                //getCellValue(row.getCell(4),evaluator).trim(); // cột E
                //getCellValue(row.getCell(5),evaluator).trim(); // cột F
                tmpi = 6;
                if (isSaoKeCaNhan) tmpi = 4;
                String date = getCellValue(row.getCell(tmpi),evaluator).trim(); // cột G
                //getCellValue(row.getCell(7),evaluator).trim(); // cột H
                tmpi = 8;
                if (isSaoKeCaNhan) tmpi = 10;
                String creditStr = getCellValue(row.getCell(tmpi),evaluator).trim(); // cột I
                //getCellValue(row.getCell(9),evaluator).trim(); // cột J
                //getCellValue(row.getCell(10),evaluator).trim(); // cột K
                //String tenhv = getCellValue(row.getCell(11),evaluator).trim(); // cột L
                tmpi = 12;
                if (isSaoKeCaNhan) tmpi = 11;
                String remark = getCellValue(row.getCell(tmpi),evaluator).trim(); // cột M
                //String remark = getCellValue(row.getCell(13),evaluator).trim(); // cột N
                //String remark = getCellValue(row.getCell(14),evaluator).trim(); // cột O
                //String remark = getCellValue(row.getCell(15),evaluator).trim(); // cột P
                //String remark = getCellValue(row.getCell(16),evaluator).trim(); // cột Q
                tmpi = 17;
                if (isSaoKeCaNhan) tmpi = 6;
                String transactionNo = getCellValue(row.getCell(tmpi),evaluator).trim(); // cột R
                
                String mahv = "";
                String tenhv = "";
                String thang = "0";

                // Nếu trống dòng → bỏ qua
                if (date.isEmpty() || transactionNo.isEmpty() || creditStr.isEmpty() || remark.isEmpty()) continue;
                
                //lay mahv
                StringUtils.RemarkInfo info = StringUtils.parseRemark(remark);
                if (info != null) {
                	mahv=info.mahv;
                	thang=String.valueOf(info.month);
                	try {
                    	monthValue = Integer.valueOf(thang);
                    } catch(Exception e1) {}
                	tenhv = info.tenhv;
                } else {
                	monthValue = 0;
                }
                
                // exists
                psCheck.setString(1, transactionNo);
                psCheck.setString(2, soTk);
                rs = psCheck.executeQuery();
                boolean exists = rs.next() && rs.getInt(1) > 0;
                rs.close();

                if (exists) {
                	
                    logMessage("Đã nhập lần trước transactionNo=" + transactionNo + ", date=" + date);
                    continue;
                }

                // amount
                double tranAmt = parseDoubleSafe(creditStr);                
                if (tranAmt==0) continue;

                // Insert
                psInsert.setString(1, soTk);
                psInsert.setInt(2, monthValue);
                psInsert.setString(3, date);
                psInsert.setString(4, transactionNo);
                psInsert.setDouble(5, tranAmt);
                psInsert.setString(6, remark);
                psInsert.setString(7, mahv);
                psInsert.setString(8, tenhv);

                psInsert.executeUpdate();
                insertedCount++;
                if (mahv.isEmpty()) insertedNoMaCount++;
            }

            workbook.close();
            fis.close();

            logMessage("Đã hoàn tất đọc sao kê. Đã thêm " + insertedCount + " dòng, trong đó " + insertedNoMaCount + " KHÔNG xác định được mã học sinh:");

            return true;
        } catch (Exception e) {
            e.printStackTrace();
            logMessage("Lỗi importBankStatement: " + e.getMessage());
            sendLogException(e);
            return false;
        } finally {
            try { if (rs != null) rs.close(); } catch (Exception e) {}
            try { if (psCheck != null) psCheck.close(); } catch (Exception e) {}
            try { if (psInsert != null) psInsert.close(); } catch (Exception e) {}
            try { if (st != null) st.close(); } catch (Exception e) {}
            try { if (conn != null) conn.close(); } catch (Exception e) {}
        }
	}
	
	public static String getStackTrace(Exception e) {
	    StringWriter sw = new StringWriter();
	    PrintWriter pw = new PrintWriter(sw);
	    e.printStackTrace(pw);
	    return sw.toString();
	}
	
	public static void addMissingColumnsToTongHopThang() {
	    Connection conn = null;
	    ResultSet rs = null;
	    try {
	        conn = DriverManager.getConnection(URL_DB);

	        // Check existing columns
	        rs = conn.createStatement().executeQuery("PRAGMA table_info(tonghopthang)");
	        java.util.Set<String> cols = new java.util.HashSet<String>();

	        while (rs.next()) {
	            cols.add(rs.getString("name").toLowerCase());
	        }
	        rs.close();

	        Statement st = conn.createStatement();

	        // Add listtranNo if missing
	        if (!cols.contains("listtranno")) {
	            st.execute("ALTER TABLE tonghopthang ADD COLUMN listtranNo TEXT");
	            System.out.println("Added column listtranNo");
	        }

	        // Add tranAmt if missing
	        if (!cols.contains("tranamt")) {
	            st.execute("ALTER TABLE tonghopthang ADD COLUMN tranAmt REAL");
	            System.out.println("Added column tranAmt");
	        }
	        //listTkThu
	        if (!cols.contains("listtkthu")) {
	            st.execute("ALTER TABLE tonghopthang ADD COLUMN listTkThu TEXT");
	            System.out.println("Added column listTkThu");
	        }
	        //phitm
	        if (!cols.contains("phitm")) {
	            st.execute("ALTER TABLE tonghopthang ADD COLUMN phitm REAL");
	            System.out.println("Added column phitm");
	        }
	        
	        st.close();

	    } catch (Exception e) {
	        e.printStackTrace();
	        sendLogException(e);
	    } finally {
	        try { if (rs != null) rs.close(); } catch (Exception ignored) {}
	        try { if (conn != null) conn.close(); } catch (Exception ignored) {}
	    }
	}

	public static void addMissingColumnsToSaoKeThang() {
	    Connection conn = null;
	    ResultSet rs = null;
	    try {
	        conn = DriverManager.getConnection(URL_DB);

	        // Check existing columns
	        rs = conn.createStatement().executeQuery("PRAGMA table_info(SaoKeThang)");
	        java.util.Set<String> cols = new java.util.HashSet<String>();

	        while (rs.next()) {
	            cols.add(rs.getString("name").toLowerCase());
	        }
	        rs.close();

	        Statement st = conn.createStatement();

	        // Add listtranNo if missing
	        if (!cols.contains("phitm")) {
	            st.execute("ALTER TABLE SaoKeThang ADD COLUMN phitm REAL");
	            System.out.println("Added column phitm");
	        }
	       
	        st.close();

	    } catch (Exception e) {
	        e.printStackTrace();
	        sendLogException(e);
	    } finally {
	        try { if (rs != null) rs.close(); } catch (Exception ignored) {}
	        try { if (conn != null) conn.close(); } catch (Exception ignored) {}
	    }
	}
	
	public void updateTongHopThangFromSaoKe() {
	    Connection conn = null;
	    PreparedStatement psGet = null;
	    PreparedStatement psUpdate = null;
	    ResultSet rs = null;

	    try {
	        conn = DriverManager.getConnection(URL_DB);

	        // Lấy toàn bộ giao dịch từ SaoKeThang
	        String sqlGet = "SELECT mahv, sotk, transactionNo, tranamt, phitm FROM SaoKeThang WHERE mahv IS NOT NULL";
	        psGet = conn.prepareStatement(sqlGet);
	        rs = psGet.executeQuery();

	        while (rs.next()) {
	            String mahv = rs.getString("mahv");
	            String sotk = rs.getString("sotk");
	            String tranNo = rs.getString("transactionNo");
	            double tranAmt = rs.getDouble("tranamt");
	            double phitm = rs.getDouble("phitm");

	            if (mahv == null || mahv.trim().isEmpty()) continue;

	            // Lấy dữ liệu hiện tại trong tonghopthang
	            String sqlSelectTH = "SELECT listtranno, listTkThu, tranAmt, phitm FROM tonghopthang WHERE upper(mahv) = ?";
	            PreparedStatement psSel = conn.prepareStatement(sqlSelectTH);
	            psSel.setString(1, mahv);
	            ResultSet rsTH = psSel.executeQuery();

	            String oldListTranNo = null;
	            String oldListTkThu = null;
	            double oldTranAmt = 0;
	            double oldphitm = 0;

	            if (rsTH.next()) {
	                oldListTranNo = rsTH.getString("listtranno");
	                oldListTkThu = rsTH.getString("listTkThu");
	                oldTranAmt = rsTH.getDouble("tranAmt");
	                oldphitm = rsTH.getDouble("phitm");
	            }
	            rsTH.close();
	            psSel.close();

	            // Xử lý listtranno
	            boolean existsTranNo = false;
	            if (oldListTranNo != null && !oldListTranNo.trim().isEmpty()) {
	                if ((","+oldListTranNo+",").contains(","+tranNo+",")){
	                	existsTranNo = true;
	                }
	            }

	            String newListTranNo = oldListTranNo;
	            double newTranAmt = oldTranAmt;
	            double newPhiTM = oldphitm;

	            if (!existsTranNo) {
	                if (newListTranNo == null || newListTranNo.trim().isEmpty()) {
	                    newListTranNo = tranNo;
	                } else {
	                    newListTranNo += "," + tranNo;
	                }
	                newTranAmt += tranAmt; // chỉ cộng khi là giao dịch mới
	                newPhiTM+=phitm;
	            }

	            // Xử lý listTkThu
	            boolean existsTk = false;
	            if (oldListTkThu != null && !oldListTkThu.trim().isEmpty()) {
	                String[] arr2 = oldListTkThu.split(",");
	                for (String s2 : arr2) {
	                    if (s2.trim().equalsIgnoreCase(sotk)) {
	                        existsTk = true;
	                        break;
	                    }
	                }
	            }

	            String newListTkThu = oldListTkThu;
	            if (!existsTk) {
	                if (newListTkThu == null || newListTkThu.trim().isEmpty()) {
	                    newListTkThu = sotk;
	                } else {
	                    newListTkThu += "," + sotk;
	                }
	            }

	            // Cập nhật tonghopthang
	            String sqlUpdate = "UPDATE tonghopthang SET listtranno = ?, listTkThu = ?, tranAmt = ?, phitm=? WHERE upper(mahv) = ?";
	            psUpdate = conn.prepareStatement(sqlUpdate);
	            psUpdate.setString(1, newListTranNo);
	            psUpdate.setString(2, newListTkThu);
	            psUpdate.setDouble(3, newTranAmt);
	            psUpdate.setDouble(4, newPhiTM);
	            psUpdate.setString(5, mahv);
	            psUpdate.executeUpdate();
	            psUpdate.close();
	        }

	        logMessage("Tổng hợp sao kê hoàn tất.");

	    } catch (Exception e) {
	        e.printStackTrace();
	        System.out.println("Lỗi updateTongHopThangFromSaoKe: " + e.getMessage());
	        sendLogException(e);
	    } finally {
	        try { if (rs != null) rs.close(); } catch (Exception e) {}
	        try { if (psGet != null) psGet.close(); } catch (Exception e) {}
	        try { if (psUpdate != null) psUpdate.close(); } catch (Exception e) {}
	        try { if (conn != null) conn.close(); } catch (Exception e) {}
	    }
	}


	    
    public void btn2ChooseFolder_Event() {
        logArea.setText("");

        // Lấy tháng từ monthSpinner
        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMM");
        int thangInt = Integer.parseInt(sdf.format((Date) monthSpinner.getValue()));

        // Thư mục mặc định theo tháng
        File rootDir = new File("BAO CAO");
        if (!rootDir.exists()) rootDir.mkdirs();
        
        File baseDir = new File(rootDir, String.valueOf(thangInt));
        if (!baseDir.exists()) baseDir.mkdirs();

        File excelDir = new File(baseDir, "excels");
        if (!excelDir.exists()) excelDir.mkdirs();
        
        JFileChooser chooser = new JFileChooser();
        chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

        // Nếu baseDir tồn tại thì đặt làm thư mục mặc định
        if (baseDir.exists()) {
            chooser.setCurrentDirectory(baseDir);
            chooser.setSelectedFile(excelDir.getAbsoluteFile()); 
        } else {
            chooser.setCurrentDirectory(rootDir.exists() ? rootDir : new File(".")); 
        }

        if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            folderPathField.setText(chooser.getSelectedFile().getAbsolutePath());
        } else {
            folderPathField.setText(baseDir.getAbsolutePath());
        }
        
        File folder = new File(chooser.getSelectedFile().getAbsolutePath());
        
        if (!folder.exists()){
            JOptionPane.showMessageDialog(this, "Thư mục không tồn tại, hãy chọn lại " + chooseButton.getText());           
        	return;
        }
        
        File[] files = folder.listFiles(new java.io.FilenameFilter() {
            public boolean accept(File dir, String name) {
                return name.toLowerCase().endsWith(".xlsx");
            }
        });
        
        int icount = 0;
        if (files != null && files.length > 0) {
            for (File file : files) {
                if (file.isFile()) {
                	icount+=1;
                    logMessage("File " + icount + ": " + file.getName());
                } else if (file.isDirectory()) {
                    //System.out.println("Thư mục: " + file.getName());
                }
            }
        }
        
        logMessage("Đã chọn thư mục ở đường dẫn: " + folderPathField.getText());
        logMessage("Tổng số files: " + files.length);
        logMessage("--------------------------------\nHãy nhấn nút: " + btnImportDataButton.getText());
        
        //importData();
    } 
    
    @SuppressWarnings("resource")
	public void btn6ExportWordFiles_Event() {
        // Thiết lập UI ban đầu
        SwingUtilities.invokeLater(new Runnable() {
            @Override
            public void run() {
                progressBar.setValue(0);
                progressBar.setString("0%");
            }
        });

        if (hasProxy) {
            System.setProperty("https.proxyHost", "10.53.120.1");
            System.setProperty("https.proxyPort", "8080");
        }

        Connection conn = null;
        PreparedStatement pstmt = null;
        ResultSet rs = null;

        try {
            SimpleDateFormat sdf = new SimpleDateFormat("yyyyMM");
            int thangInt = Integer.parseInt(sdf.format((Date) monthSpinner.getValue()));
            
          //kiem tra chua co du lieu tong hop thang thi quit
        	if (!checkTranDataExists(thangInt, false)) {
        		logMessage("Chưa có dữ liệu tổng hợp tháng!");
        		logMessage("=> Hãy chọn " + btnImportDataButton.getText() + " hoặc " + btnImportNHButton.getText());
        		JOptionPane.showMessageDialog(this, "Chưa có dữ liệu tổng hợp tháng!");
        		return;
        	}
            
            SimpleDateFormat sdfText = new SimpleDateFormat("MM-yyyy");
            String sThang = sdfText.format((Date) monthSpinner.getValue());
            SimpleDateFormat sdfTextThang = new SimpleDateFormat("MM");
            String sOnlyMonth = sdfTextThang.format((Date) monthSpinner.getValue());

            // 1. Tạo thư mục
            File baocaoDir = new File("BAO CAO");
            if (!baocaoDir.exists()) baocaoDir.mkdirs();

            File baseDir = new File(baocaoDir, String.valueOf(thangInt));
            if (!baseDir.exists()) baseDir.mkdirs();

            File qrDir = new File(baseDir, "qr");
            if (qrDir.exists()) {
            	deleteDirectory(qrDir);
            }
            qrDir.mkdirs();

            File hvDir = new File(baseDir, "hv");
            if (hvDir.exists()) {
            	deleteDirectory(hvDir);
            }
            hvDir.mkdirs();
            
            File wordsDir = new File(baseDir, "words");
            if (wordsDir.exists()) {
            	deleteDirectory(wordsDir);
            }
            wordsDir.mkdirs();
            
            deleteDocxFiles(baseDir);

            logMessage("Bắt đầu tạo file cho từng học sinh ...",true);
            
            // 2. Kết nối DB và đếm tổng số bản ghi
            conn = DriverManager.getConnection(URL_DB);
            pstmt = conn.prepareStatement("SELECT COUNT(*) FROM tonghopthang");
            rs = pstmt.executeQuery();
            int totalRecords = rs.next() ? rs.getInt(1) : 0;
            rs.close();

            final int finalTotalRecords = totalRecords;
            SwingUtilities.invokeLater(new Runnable() {
                @Override
                public void run() {
                    progressBar.setMaximum(finalTotalRecords + 1);
                }
            });

            pstmt = conn.prepareStatement("SELECT id, thang, tenhv, mahv, loptt, monToan, monVan, monAnh, monLy, monHoa, sotk, tentk, tennh, dongia, tongsongay FROM tonghopthang");
            rs = pstmt.executeQuery();

            final java.util.concurrent.atomic.AtomicInteger processedRecords = new java.util.concurrent.atomic.AtomicInteger(0);

            while (rs.next()) {
                String tenhv = rs.getString("tenhv");
                String mahv = rs.getString("mahv");
                String sotk = rs.getString("sotk");
                String tentk = rs.getString("tentk");
                String tennh = rs.getString("tennh");
                String loptt = rs.getString("loptt");

                String monToan = rs.getString("monToan");
                String monVan = rs.getString("monVan");
                String monAnh = rs.getString("monAnh");
                String monLy = rs.getString("monLy");
                String monHoa = rs.getString("monHoa");
                String sTongSoNgay = rs.getString("tongsongay");

                double dongia = rs.getDouble("dongia");
                double tongsongay = rs.getDouble("tongsongay");
                double tongtien = dongia * tongsongay;

                if (sotk == null || sotk.isEmpty()) throw new Exception("Thiếu số tài khoản NH hãy xem lại bước " + btnImportNHButton.getText());
                
                String des = removeVietnameseAccents(tenhv + " " + mahv + " thang " + sOnlyMonth);
                String encodedDes = URLEncoder.encode(des, java.nio.charset.StandardCharsets.UTF_8.name());

                String url = "https://qr.sepay.vn/img?bank=MB&acc=" + sotk +
                        "&template=qronly&amount=" + String.valueOf(tongtien) +
                        "&download=DOWNLOAD&des=" + encodedDes;

                //System.out.println(url);

                String imagePath = downloadQrCodeWithExt(url, qrDir.getAbsolutePath(), mahv);
                
                File imageFile = new File(imagePath);
                if (!imageFile.exists()) throw new Exception("Không tạo được mã QR! Hãy kiểm tra lại file Tài khoản NH và làm lại " + btnImportNHButton.getText());

                Map<String, String> vars = new java.util.HashMap<String, String>();
                vars.put("thang", sThang);
                vars.put("tenhv", tenhv);
                vars.put("mahv", mahv);
                vars.put("sotk", sotk != null ? sotk : "");
                vars.put("tentk", tentk != null ? tentk : "");
                vars.put("tennh", tennh != null ? tennh : "");
                vars.put("noidung", des);
                vars.put("loptt", loptt != null ? loptt : "");
                vars.put("math", monToan != null ? monToan : "");
                vars.put("van", monVan != null ? monVan : "");
                vars.put("eng", monAnh != null ? monAnh : "");
                vars.put("phy", monLy != null ? monLy : "");
                vars.put("chem", monHoa != null ? monHoa : "");
                vars.put("price", formatCurrency(String.valueOf(dongia)));
                vars.put("total", sTongSoNgay != null ? sTongSoNgay : "");
                vars.put("amt", formatCurrency(String.valueOf(tongtien)));

                String outputPath = new File(hvDir, loptt + "_" + mahv + ".docx").getAbsolutePath();   
                
                WordTemplate.generateWordFromTemplate(TEMPLATE_DOCX, outputPath, vars, "");

                logMessage("Đã tạo file Word cho: " + tenhv + " (" + mahv + ")");

                final int currentProgress = processedRecords.incrementAndGet();
                SwingUtilities.invokeLater(new Runnable() {
                    @Override
                    public void run() {
                        progressBar.setValue(currentProgress);
                        progressBar.setString(String.format("%d/%d Học sinh (%d%%)",
                                currentProgress,
                                finalTotalRecords,
                                (currentProgress * 100 / finalTotalRecords)));
                    }
                });
            }

            mergeDocxFiles(hvDir.getAbsolutePath(), wordsDir.getAbsolutePath(), qrDir.getAbsolutePath());
                       
            
            SwingUtilities.invokeLater(new Runnable() {
                @Override
                public void run() {
                    progressBar.setValue(finalTotalRecords + 1);
                    progressBar.setString("Hoàn tất!");
                }
            });

            SwingUtilities.invokeLater(new Runnable() {
                @Override
                public void run() {
                    
                    progressBar.setValue(0);
                    progressBar.setString("Sẵn sàng");
                }
            });
            
            logMessage("Xuất file Word hoàn tất. Thư mục: " + wordsDir.getAbsolutePath());            
            JOptionPane.showMessageDialog(this, "Xuất file Word hoàn tất. Thư mục: " + wordsDir.getAbsolutePath());

            DocxMerger.openBaseDir(wordsDir.getAbsolutePath());
        } catch (Exception e) {
            e.printStackTrace();
            logMessage("Lỗi exportWordFiles: " + e.getMessage());
            SwingUtilities.invokeLater(new Runnable() {
                @Override
                public void run() {
                    progressBar.setValue(0);
                    progressBar.setString("Lỗi");
                }
            });
            sendLogException(e);
        } finally {
            try { if (rs != null) rs.close(); } catch (Exception e) { }
            try { if (pstmt != null) pstmt.close(); } catch (Exception e) { }
            try { if (conn != null) conn.close(); } catch (Exception e) { }
        }
    }

    @SuppressWarnings("resource")
	public void mergeDocxFiles(String inputFolderPath, String outputFolderPath, String imageDirPath) throws IOException, XmlException, InvalidFormatException {
    	logMessage("Bắt đầu tạo file word cho lớp... ", true);    	
    	
    	Path inputDir = Paths.get(inputFolderPath);
        Path outputDir = Paths.get(outputFolderPath);

        if (!Files.exists(inputDir) || !Files.isDirectory(inputDir)) {
        	logMessage("Input folder does not exist or is not a directory: " + inputFolderPath);
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
        	logMessage("Input folder is empty or access denied: " + inputFolderPath);
        }

        final int finalTotalRecords = filesToMerge.size();
        SwingUtilities.invokeLater(new Runnable() {
            @Override
            public void run() {
                progressBar.setMaximum(finalTotalRecords + 1);
            }
        });
        
        final java.util.concurrent.atomic.AtomicInteger processedRecords = new java.util.concurrent.atomic.AtomicInteger(0);
        for (Map.Entry<String, List<Path>> entry : filesToMerge.entrySet()) {
            String baseName = entry.getKey();
            List<Path> docxFiles = entry.getValue();

            	
                XWPFDocument mergedDocument = null; // Khởi tạo null để handle đóng tài liệu
               
                try {
                	String mahv0=HVMain.extractCodeFromPath(docxFiles.get(0).toFile().getAbsolutePath());                             
                    File imagePath0 = new File(imageDirPath,mahv0+".png");
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
                            File imagePath = new File(imageDirPath,mahv+".png");
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
                } finally {
                    // Đảm bảo đóng mergedDocument nếu nó đã được tạo
                    if (mergedDocument != null) {
                        try {
                            mergedDocument.close();
                        } catch (IOException e) {
                            System.err.println("Error closing merged document for " + baseName + ": " + e.getMessage());
                            e.printStackTrace();
                        }
                    }
                }
            
            logMessage("Đã tạo file word cho lớp: " + baseName);
            
            final int currentProgress = processedRecords.incrementAndGet();
            SwingUtilities.invokeLater(new Runnable() {
                @Override
                public void run() {
                    progressBar.setValue(currentProgress);
                    progressBar.setString(String.format("%d/%d lớp (%d%%)",
                            currentProgress,
                            finalTotalRecords,
                            (currentProgress * 100 / finalTotalRecords)));
                }
            });
        }
        
        
        
    }
    
    public boolean deleteDirectory(File dir) {
        if (dir == null || !dir.exists()) {
            return false;
        }

        if (dir.isDirectory()) {
            File[] files = dir.listFiles();
            if (files != null) {
                for (File f : files) {
                    deleteDirectory(f); // xóa file hoặc thư mục con
                }
            }
        }

        return dir.delete(); // xóa thư mục hoặc file cuối cùng
    }
    
    public void deleteDocxFiles(File baseDir) {
        if (baseDir == null || !baseDir.exists() || !baseDir.isDirectory()) {
            logMessage("Thư mục baseDir không hợp lệ!");
            return;
        }

        File[] files = baseDir.listFiles(new FilenameFilter() {
            @Override
            public boolean accept(File dir, String name) {
                return name.toLowerCase().endsWith(".docx");
            }
        });

        if (files == null || files.length == 0) {
            logMessage("Không tìm thấy file .docx trong thư mục: " + baseDir.getAbsolutePath());
            return;
        }

        for (File file : files) {
            if (file.delete()) {
                logMessage("Đã xóa: " + file.getName());
            } else {
                logMessage("Không thể xóa: " + file.getName());
            }
        }
    }

    public static String formatCurrency(String amountStr) {
        if (amountStr == null || amountStr.trim().isEmpty()) {
            return "";
        }
        try {
            double value = Double.parseDouble(amountStr);
            DecimalFormat formatter = new DecimalFormat("#,###");
            return formatter.format(value);
        } catch (NumberFormatException e) {
            // Nếu đầu vào không phải số, trả về nguyên chuỗi
            return amountStr;
        }
    }
    
   
 // Tải QR code và trả về đường dẫn file với đuôi đúng
    public static String downloadQrCodeWithExt(String urlStr, String outputDir, String fileBaseName) throws Exception {
    	if (hasProxy) {
            System.setProperty("https.proxyHost", "10.53.120.1");
            System.setProperty("https.proxyPort", "8080");
        }
    	
    	URL url = new URL(urlStr);
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setRequestProperty("User-Agent", "Mozilla/5.0");
        conn.connect();

        if (conn.getResponseCode() != 200) {
            throw new RuntimeException("Lỗi tải QR: " + conn.getResponseCode() + " " + conn.getResponseMessage());
        }

        String contentType = conn.getContentType(); // ví dụ: image/png, image/jpeg
        String ext = ".jpg";  // mặc định
        if (contentType != null) {
            if (contentType.contains("png")) ext = ".png";
            else if (contentType.contains("jpeg") || contentType.contains("jpg")) ext = ".jpg";
        }

        String fileName = fileBaseName + ext;
        File outFile = new File(outputDir, fileName);

        InputStream in = conn.getInputStream();
        FileOutputStream out = new FileOutputStream(outFile);
        byte[] buffer = new byte[4096];
        int bytesRead;
        while ((bytesRead = in.read(buffer)) != -1) {
            out.write(buffer, 0, bytesRead);
        }
        out.close();
        in.close();
        conn.disconnect();

        return outFile.getAbsolutePath(); // path đầy đủ
    }
    
    public static String removeVietnameseAccents(String input) {
	    if (input == null) return null;

	    // Chuẩn hóa chuỗi thành dạng tổ hợp ký tự (NFD)
	    String normalized = Normalizer.normalize(input, Normalizer.Form.NFD);

	    // Bỏ toàn bộ dấu (ký tự có thuộc tính "Mark")
	    String withoutAccents = normalized.replaceAll("\\p{M}", "");

	    // Chuẩn hóa một số ký tự đặc biệt còn sót
	    withoutAccents = withoutAccents.replaceAll("Đ", "D").replaceAll("đ", "d");

	    return withoutAccents;
	}
    
    public boolean isValidMonHoc(String monhoc) {
        if (monhoc == null) return false;
        return monhoc.equals("Toán") ||
               monhoc.equals("Ngữ Văn") ||
               monhoc.equals("Tiếng Anh") ||
               monhoc.equals("Lý") ||
               monhoc.equals("Hóa");
    }
    
    public void btn3ImportData_Event() {
    	SimpleDateFormat sdf = new SimpleDateFormat("yyyyMM");
        int thang = Integer.parseInt(sdf.format((Date) monthSpinner.getValue()));
        
        String folderPath = folderPathField.getText();
        if (folderPath.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Chưa chọn thư mục!");
            return;
        }

        File folder = new File(folderPath);
        File[] files = folder.listFiles(new java.io.FilenameFilter() {
            public boolean accept(File dir, String name) {
                return name.toLowerCase().endsWith(".xlsx");
            }
        });

        if (files == null || files.length == 0) {
            JOptionPane.showMessageDialog(this, "Không có file Excel trong thư mục!");
            return;
        }

        //check data
        if (ConfirmNotOverideExistData(thang)) {
        	logMessage("Đã có dữ liệu tổng hợp tháng!", true);
        	return;
        }
    	//end check data
        
        
        progressBar.setMaximum(files.length);
        progressBar.setValue(0);
        logArea.setText("Đã chọn thư mục ở đường dẫn: " + folderPathField.getText() + "\n");

        Connection conn = null;
        PreparedStatement ps = null;
        try {
            conn = DriverManager.getConnection(URL_DB);
            Statement st = conn.createStatement();
            
            st.execute("DROP TABLE IF EXISTS tonghop");
            st.execute("DROP TABLE IF EXISTS tonghopthang");
            st.execute("DROP TABLE IF EXISTS SaoKeThang");
            st.execute("CREATE TABLE IF NOT EXISTS tonghop (id INTEGER PRIMARY KEY AUTOINCREMENT, thang INTEGER, tenhv TEXT, loptt TEXT, songay REAL, lophv TEXT, tengv TEXT, monhoc TEXT, dongia REAL, ghichu TEXT, mahv TEXT, magv TEXT, file TEXT, thanhtien REAL)");
            //PreparedStatement psDelete = conn.prepareStatement("DELETE FROM tonghop WHERE thang = ?");
        	//psDelete.setInt(1, thang);
        	//psDelete.executeUpdate();

            String sql = "INSERT INTO tonghop(thang, tenhv, loptt, songay, lophv, tengv, monhoc, dongia, ghichu, mahv, magv, file, thanhtien) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";
            ps = conn.prepareStatement(sql);

            

            int fileCount = 0;
            int totalImported = 0;
            double totalAmt = 0;
            for (File file : files) {
                FileInputStream fis = null;
                Workbook workbook = null;
                int rowCount = 0;
                try {
                    fis = new FileInputStream(file);
                    workbook = new XSSFWorkbook(fis);
                    Sheet sheet = workbook.getSheetAt(0);

                    FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                    
                    String lophv = sheet.getRow(3).getCell(2).toString().trim();
                    String tengv = sheet.getRow(4).getCell(4).toString().trim();
                    String monhoc = sheet.getRow(4).getCell(2).toString().trim();

                    if (!isValidMonHoc(monhoc)) {
                    	logMessage("Yêu cầu nhập tên môn học là: Toán, Ngữ Văn, Tiếng Anh, Lý, Hóa");
                    	logMessage("Hãy sửa lại file " + file.getName() + " và thực hiện lại " + btnImportDataButton.getText());
                    	return;
                    }
                    
                    for (int i = 9; ; i++) {
                        Row row = sheet.getRow(i);
                        if (row == null || row.getCell(0) == null || row.getCell(0).toString().trim().isEmpty()) {
                            break;
                        }

                        //String stt = getCellValue(row.getCell(0));
                        String tenhv = getCellValue(row.getCell(1));
                        if (tenhv.isEmpty()) continue;
                        
                       

                        String loptt = getCellValue(row.getCell(2), evaluator).toUpperCase();
                        String ghichu = getCellValue(row.getCell(4), evaluator);
                        String songayStr = getCellValue(row.getCell(5), evaluator);
                        String dongiaStr = getCellValue(row.getCell(6), evaluator);

                        if (test & loptt.equalsIgnoreCase("C10")){
                        	System.out.println("C10 in " + file.getAbsolutePath());
                        }
                        //System.out.println(stt + "-" + tenhv + "-" + loptt + "-" + songayStr);
                        
                        double songay = 0;
                        if (!songayStr.isEmpty()) {
                            try {
                                songay = Double.parseDouble(songayStr.replaceAll("[^0-9.]", ""));
                            } catch (NumberFormatException e) {
                                songay = 0;
                            }
                        }

                        double dongia = 0;
                        if (!dongiaStr.isEmpty()) {
                            try {
                                dongia = Double.parseDouble(dongiaStr.replaceAll("[^0-9.]", ""));
                            } catch (NumberFormatException e) {
                                dongia = 0;
                            }
                        }

                        double thanhtien = songay * dongia;

                        ps.setInt(1, thang);
                        ps.setString(2, tenhv);
                        ps.setString(3, loptt);
                        ps.setDouble(4, songay);
                        ps.setString(5, lophv);
                        ps.setString(6, tengv);
                        ps.setString(7, monhoc);
                        ps.setDouble(8, dongia);
                        ps.setString(9, ghichu);
                        ps.setString(10, null);
                        ps.setString(11, null);
                        ps.setString(12, file.getName());
                        ps.setDouble(13, thanhtien);

                        //validate:
                        
                        
                        
                        ps.executeUpdate();
                        rowCount++;
                        totalImported++;
                        totalAmt=totalAmt+thanhtien;
                    }

                    String logMsg = "File " + file.getName() + ": Đã import " + rowCount + " dòng.";
                    logMessage(logMsg);
                } finally {
                    if (workbook != null) workbook.close();
                    if (fis != null) fis.close();
                }
                fileCount++;
                final int progress = fileCount;
                final int totalFile = files.length;
                final int totalRecords = totalImported;
                final String totalAmount = formatCurrency(String.valueOf(totalAmt));
                SwingUtilities.invokeLater(new Runnable() {
                    public void run() {                        
                        progressBar.setValue(progress);
                        progressBar.setString(String.format("%d/%d file (%d%%) %d dòng " + totalAmount + "VND",
                        		progress,
                        		totalFile,
                                (progress * 100 / totalFile),
                                totalRecords));
                    }
                });
            }
            
            // Xuất file Excel tổng hợp
            String fileName=exportClassExcelToEntryBank(conn, thang);
            
            createHV();
            
            fixOnlyDuplicateMahv();
            
            updateMahvToTonghop();
            
            String summary = "Tổng cộng đã import " + totalImported + " dòng từ " + files.length + " files.";            
            logMessage(summary);
            logMessage("Tổng tiền " + formatCurrency(String.valueOf(totalAmt)));
            logMessage("---------------------------------");
            logMessage("Nhập tài khoản NH vào file excel: " + fileName);
            logMessage("=> Sau đó chọn " + btnImportNHButton.getText() + " để chọn file đã nhập!");
            
            JOptionPane.showMessageDialog(this, summary);
            
            //DocxMerger.openBaseDir(baseDir.getAbsolutePath());
            DocxMerger.openDesktop(fileName);
        } catch (Exception ex) {
            ex.printStackTrace();
            logMessage("Lỗi: " + ex.getMessage());
            sendLogException(ex);
            JOptionPane.showMessageDialog(this, "Lỗi: " + ex.getMessage());
        } finally {
            try {
                if (ps != null) ps.close();
                if (conn != null) conn.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }


	public boolean ConfirmNotOverideExistData(int thang) {
		String message = "Đã có dữ liệu tổng hợp tháng " + thang + "! Bạn có muốn chạy lại không?";
        
		boolean checkExistPayment = checkTranDataExists(thang,true);
        boolean checkExist = false;
        if (!checkExistPayment){
        	checkExist = checkTranDataExists(thang,false);
        } else {
        	message = "Dữ liệu tháng " + thang + " đang được thanh toán! Bạn có muốn chạy lại không?";
        }
        
    	if (checkExist || checkExistPayment){
		   int choice = JOptionPane.showConfirmDialog(
		           this,
		           message,
		           "Xác nhận",
		           JOptionPane.YES_NO_OPTION,
		           JOptionPane.WARNING_MESSAGE
		   );
		
		   if (choice != JOptionPane.YES_OPTION) return true;
		   
		   HtmlReader.BackupOnly();
    	}
    	return false;
	}
    
    public String btn5ExportToSumaryExcel_Event() {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMM");
        int thangInt = Integer.parseInt(sdf.format((Date) monthSpinner.getValue()));

        //kiem tra chua co du lieu tong hop thang thi quit
    	if (!checkTranDataExists(thangInt, false)) {
    		logMessage("Chưa có dữ liệu tổng hợp tháng!");
    		logMessage("=> Hãy chọn " + btnImportDataButton.getText() + " hoặc " + btnImportNHButton.getText());
    		JOptionPane.showMessageDialog(this, "Chưa có dữ liệu tổng hợp tháng!");
    		return "";
    	}
        
        File rootDir = new File("BAO CAO");
        if (!rootDir.exists()) rootDir.mkdirs();

        File baseDir = new File(rootDir, String.valueOf(thangInt));
        if (!baseDir.exists()) baseDir.mkdirs();

        addMissingColumnsToTongHopThang();
        
        Connection conn = null;
        PreparedStatement pstmt = null;
        ResultSet rs = null;

        try {
            conn = DriverManager.getConnection(URL_DB);

            // 1. Đếm tổng số bản ghi để set progressBar
            int totalRecords = 0;
            pstmt = conn.prepareStatement("SELECT COUNT(*) FROM tonghopthang");
            rs = pstmt.executeQuery();
            if (rs.next()) totalRecords = rs.getInt(1);
            rs.close();
            pstmt.close();

            final int finalTotal = totalRecords;
            SwingUtilities.invokeLater(new Runnable() {
                @Override
                public void run() {
                    progressBar.setMinimum(0);
                    progressBar.setMaximum(finalTotal);
                    progressBar.setValue(0);
                    progressBar.setString("0%");
                }
            });

            // 2. Lấy dữ liệu
            String sql = "SELECT id, thang, tenhv, mahv, loptt, monToan, monVan, monAnh, " +
                         "monLy, monHoa, dongia, tongsongay, dongia*tongsongay as tongtien, "
                         + "tranAmt,phitm, dongia*tongsongay-IFNULL(CAST(tranAmt AS REAL), 0.0) as chenhlech,listTkThu,listtranNo," +
                         "sotk, tentk, tennh FROM tonghopthang";
            pstmt = conn.prepareStatement(sql);
            rs = pstmt.executeQuery();

            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("BaoCao" + thangInt);

            // Header
            String[] headers = {
                "ID", "Tháng", "Tên HV", "Mã HV", "Lớp TT",
                "Môn Toán", "Môn Văn", "Môn Anh", "Môn Lý", "Môn Hóa",
                "Đơn giá", "Tổng số lượt", " Tổng số tiền ",
                "Số thu được","CT","TH","TM","Phí TM","Đóng dư tiền",
                "Số TK", "Tên TK", "Tên NH"
            };

            Row headerRow = sheet.createRow(0);
            CellStyle headerStyle = workbook.createCellStyle();
            org.apache.poi.ss.usermodel.Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerStyle.setFont(headerFont);

            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
                cell.setCellStyle(headerStyle);
            }

            // 3. Ghi dữ liệu
            
            int tSoHSChuaDong = 0;
            double tSoTienChuaDong = 0.0;
            int tSoHS =0;
            double tSoTienPhaiThu=0.0;
            double tSoTienThuDuoc=0.0;
            //double tSoTienChuaThuDuoc = 0.0;
            double tThuTKCongTy =0.0;
            double tThuTKThuHo = 0.0;
            double tThuTM = 0.0;
            double tPhiTM = 0.0;
            
            int rowNum = 1;
            int processed = 0;
            while (rs.next()) {
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(rs.getInt("id"));
                row.createCell(1).setCellValue(rs.getInt("thang"));
                row.createCell(2).setCellValue(rs.getString("tenhv"));
                row.createCell(3).setCellValue(rs.getString("mahv"));
                row.createCell(4).setCellValue(rs.getString("loptt"));
                row.createCell(5).setCellValue(rs.getDouble("monToan"));
                row.createCell(6).setCellValue(rs.getDouble("monVan"));
                row.createCell(7).setCellValue(rs.getDouble("monAnh"));
                row.createCell(8).setCellValue(rs.getDouble("monLy"));
                row.createCell(9).setCellValue(rs.getDouble("monHoa"));
                row.createCell(10).setCellValue(rs.getDouble("dongia"));
                row.createCell(11).setCellValue(rs.getDouble("tongsongay"));
                row.createCell(12).setCellValue(rs.getDouble("tongtien"));
                
                row.createCell(13).setCellValue(rs.getDouble("tranAmt"));
                
                String tkThu = rs.getString("listTkThu");
                if (tkThu == null || tkThu.isEmpty()) tkThu="";
                String CT = "", TH = "", TM="";
                String listtranNo = rs.getString("listtranNo");
                
                if (listtranNo == null || listtranNo.isEmpty()) listtranNo="";
                if ((","+tkThu+",").contains(",2588666868,")) {
                	CT=listtranNo;
                	tThuTKCongTy += rs.getDouble("tranAmt");
                }
                else {
                	if ((","+tkThu+",").contains(",TM,")) {
                		TM=listtranNo;
                		tThuTM += rs.getDouble("tranAmt");
                		tPhiTM += rs.getDouble("phitm");
                    } else {
                    	TH=listtranNo;
                    	tThuTKThuHo+=rs.getDouble("tranAmt");
                    }
                }
                
                row.createCell(14).setCellValue(CT);//CT
                row.createCell(15).setCellValue(TH);//TH
                row.createCell(16).setCellValue(TM);//TM
                row.createCell(17).setCellValue(rs.getDouble("phitm"));//PhiTM
                row.createCell(18).setCellValue(rs.getDouble("chenhlech"));
                row.createCell(19).setCellValue(rs.getString("sotk"));
                row.createCell(20).setCellValue(rs.getString("tentk"));
                row.createCell(21).setCellValue(rs.getString("tennh"));

                
                if (rs.getDouble("chenhlech")==rs.getDouble("tongtien")){
                	tSoHSChuaDong+=1;
                	tSoTienChuaDong+=rs.getDouble("tongtien");
                }
                tSoHS = rowNum;
                tSoTienPhaiThu += rs.getDouble("tongtien");
                tSoTienThuDuoc += rs.getDouble("tranAmt");
                
                processed++;
                final int progressCopy = processed;
                SwingUtilities.invokeLater(new Runnable() {
                    @Override
                    public void run() {
                        progressBar.setValue(progressCopy);
                        progressBar.setString(
                            progressCopy + "/" + finalTotal + " (" + (progressCopy * 100 / finalTotal) + "%)"
                        );
                    }
                });
            }
            
            sheet.createRow(rowNum++);
            
            Row row1 = sheet.createRow(rowNum++);
            row1.createCell(2).setCellValue("Số học sinh chưa đóng");
            row1.createCell(3).setCellValue(tSoHSChuaDong);
            
            Row row2 = sheet.createRow(rowNum++);
            row2.createCell(2).setCellValue("Tương ứng số tiền");
            row2.createCell(3).setCellValue(tSoTienChuaDong);
            
            Row row3 = sheet.createRow(rowNum++);
            row3.createCell(2).setCellValue("Tổng số học sinh");
            row3.createCell(3).setCellValue(tSoHS-1);
            
            Row row4 = sheet.createRow(rowNum++);
            row4.createCell(2).setCellValue("Số tiền phải thu");
            row4.createCell(3).setCellValue(tSoTienPhaiThu);
            
            Row row5 = sheet.createRow(rowNum++);
            row5.createCell(2).setCellValue("Số tiền thu được");
            row5.createCell(3).setCellValue(tSoTienThuDuoc);
            
            Row row6 = sheet.createRow(rowNum++);
            row6.createCell(2).setCellValue("Tỷ lệ thu (%)");
            row6.createCell(3).setCellValue(Double.valueOf(tSoTienThuDuoc*100/tSoTienPhaiThu));
            
            Row row7 = sheet.createRow(rowNum++);
            row7.createCell(2).setCellValue("Thu tk công ty");
            row7.createCell(3).setCellValue(tThuTKCongTy);
            
            Row row8 = sheet.createRow(rowNum++);
            row8.createCell(2).setCellValue("Thu tk thu hộ");
            row8.createCell(3).setCellValue(tThuTKThuHo);
            
            Row row9 = sheet.createRow(rowNum++);
            row9.createCell(2).setCellValue("Thu tiền mặt");
            row9.createCell(3).setCellValue(tThuTM);
            
            Row row10 = sheet.createRow(rowNum++);
            row10.createCell(2).setCellValue("Chi trả thu tiền mặt");
            row10.createCell(3).setCellValue(tPhiTM);

            // Auto fit cột
            for (int i = 0; i < headers.length; i++) {
                sheet.autoSizeColumn(i);
            }

            // 4. Xuất file Excel
            LocalDateTime now = LocalDateTime.now();
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd_HHmmss");
            String formattedDateTime = now.format(formatter);
            
            File outputFile = new File(baseDir, "Tong hop " + thangInt + formattedDateTime +".xlsx");
            FileOutputStream fos = new FileOutputStream(outputFile);
            workbook.write(fos);
            fos.close();
            workbook.close();

            SwingUtilities.invokeLater(new Runnable() {
                @Override
                public void run() {
                    progressBar.setString("Hoàn tất!");
                }
            });

            logMessage("Đã tạo file excel tổng hợp : " + outputFile.getAbsolutePath());
            
            //DocxMerger.openBaseDir(baseDir.getAbsolutePath());
            
            return outputFile.getAbsolutePath();

        } catch (Exception e) {
            e.printStackTrace();
            if (e.toString().contains("[SQLITE_ERROR] SQL error or missing database (no such table: tonghopthang)")) {
            	logMessage("Cần thực hiện bước " + btnImportNHButton.getText());
            	return "";
            } else if (e.toString().contains("(The process cannot access the file because it is being used by another process)")){
            	logMessage("File excel tổng hợp đang mở, hãy tắt đi và xuất lại bằng " + btnExportExcelButton.getText());
            	return "";
            }
            SwingUtilities.invokeLater(new Runnable() {
                @Override
                public void run() {
                    progressBar.setValue(0);
                    progressBar.setString("Lỗi!");
                }
            });
            sendLogException(e);
            return "";
        } finally {
            try { if (rs != null) rs.close(); } catch (Exception ignored) {}
            try { if (pstmt != null) pstmt.close(); } catch (Exception ignored) {}
            try { if (conn != null) conn.close(); } catch (Exception ignored) {}
        }
    }
    
    public String exportClassExcelToEntryBank(Connection conn, int thang) throws Exception {
    	 File rootDir = new File("BAO CAO");
         if (!rootDir.exists()) rootDir.mkdirs();
         
         File baseDir = new File(rootDir, String.valueOf(thang));
         if (!baseDir.exists()) baseDir.mkdirs();
         else baseDir.delete();

         File qrDir = new File(baseDir, "qr");
         if (!qrDir.exists()) qrDir.mkdirs();
         
         File hvDir = new File(baseDir, "hv");
         if (!hvDir.exists()) hvDir.mkdirs();
    	
    	
    	String sql = "SELECT loptt, SUM(thanhtien) as tongtien FROM tonghop GROUP BY loptt";
        Statement st = conn.createStatement();
        ResultSet rs = st.executeQuery(sql);

        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Tong hop");

        // Tiêu đề
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("STT");
        header.createCell(1).setCellValue("Lớp TT");
        header.createCell(2).setCellValue("Tổng tiền");
        header.createCell(3).setCellValue("Số TK");
        header.createCell(4).setCellValue("Tên TK");
        header.createCell(5).setCellValue("Tên Ngân hàng");

        int rowIndex = 1;
        int stt = 1;
        while (rs.next()) {
            Row row = sheet.createRow(rowIndex++);
            row.createCell(0).setCellValue(stt++);
            row.createCell(1).setCellValue(rs.getString("loptt"));
            row.createCell(2).setCellValue(rs.getDouble("tongtien"));
            row.createCell(3).setCellValue("");
            row.createCell(4).setCellValue("");
            row.createCell(5).setCellValue("Ngân hàng MB");
        }

        //java.text.SimpleDateFormat sdfx = new java.text.SimpleDateFormat("yyyy.MM.dd_HH.mm");
        //String timestamp = sdfx.format(new java.util.Date());

        // Đặt tên file theo yêu cầu
        String fileName = "Nhap tai khoan NH_" + thang + ".xlsx";
        File outFile = new File(baseDir,fileName);
        FileOutputStream fos = new FileOutputStream(outFile);
        
        wb.write(fos);
        fos.close();
        wb.close();
        rs.close();
        st.close();
        
        return (outFile.getAbsolutePath());
    }

    public String replaceMiddleCharWithX(String input) {
        if (input == null || input.length() < 3) {
            return input; // nếu quá ngắn thì giữ nguyên
        }

        // tìm ký tự đầu là số, ký tự cuối là số → thay đoạn giữa = "x"
        char first = input.charAt(0);
        char last = input.charAt(input.length() - 1);

        if (Character.isDigit(first) && Character.isDigit(last)) {
            // ví dụ "10A7" → "10x7"
            return input.replaceAll("(?<=\\d)[A-Za-z]+(?=\\d)", "x");
        }

        // nếu không theo định dạng này, giữ nguyên
        return input;
    }
    
    public void fixOnlyDuplicateMahv() {
        Connection conn = null;
        PreparedStatement psFindDup = null;
        PreparedStatement psSelectGroup = null;
        PreparedStatement psUpdate = null;

        try {
            conn = DriverManager.getConnection(URL_DB);

            // 1. Lấy danh sách mahv bị trùng
            psFindDup = conn.prepareStatement(
                "SELECT mahv FROM hv GROUP BY mahv HAVING COUNT(*) > 1"
            );
            
            psUpdate = conn.prepareStatement(
			        "UPDATE hv SET mahv=? WHERE id=?"
			);

            ResultSet rsDup = psFindDup.executeQuery();
            

            while (rsDup.next()) {
                String duplicatedMahv = rsDup.getString("mahv");

                // 2. Lấy nhóm trùng theo id tăng dần
                psSelectGroup = conn.prepareStatement(
                    "SELECT id, tenhv, loptt FROM hv WHERE mahv=? ORDER BY id"
                );
                psSelectGroup.setString(1, duplicatedMahv);

                ResultSet rsGroup = psSelectGroup.executeQuery();

                java.util.List<Integer> ids = new java.util.ArrayList<Integer>();
                java.util.List<String> loptts = new java.util.ArrayList<String>();

                while (rsGroup.next()) {
                    ids.add(rsGroup.getInt("id"));
                    loptts.add(rsGroup.getString("loptt"));
                }
                rsGroup.close();

                if (ids.size() <= 1) continue; // chỉ trùng mới xử lý

                // --- Dòng đầu tiên giữ nguyên ---
                //int idKeep = ids.get(0);

                // --- Các bản ghi sau → sửa mã ---
                for (int i = 1; i < ids.size(); i++) {
                    int id = ids.get(i);
                    String loptt = loptts.get(i);

                    // Mỗi lớp tự sinh chỉ số chạy tiếp theo
                    int myIdx = getNextIndexForClass(loptt);

                    String malop = replaceMiddleCharWithX(loptt);
                    String newMahv = "SV" + malop + "x" + myIdx;

                   
        			psUpdate.setString(1, newMahv);
        			psUpdate.setInt(2, id);
        			psUpdate.executeUpdate();
        			

                    logMessage("Sửa trùng mã ID=" + id + " → " + newMahv);
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
            logMessage("Lỗi fixOnlyDuplicateMahv: " + e.getMessage());
        } finally {
            try { if (psFindDup != null) psFindDup.close(); } catch (Exception e) {}
            try { if (psSelectGroup != null) psSelectGroup.close(); } catch (Exception e) {}
            try { if (psUpdate != null) psUpdate.close(); } catch (Exception e) {}
            try { if (conn != null) conn.close(); } catch (Exception e) {}
        }
    }

    public static int getNextIndexForClass(String loptt) {
    	Connection conn = null;
    	int maxIndex = 0;
        try {
            conn = DriverManager.getConnection(URL_DB);
            PreparedStatement ps = conn.prepareStatement(
                    "SELECT mahv FROM hv WHERE loptt=? AND mahv IS NOT NULL"
                );
                ps.setString(1, loptt);
                ResultSet rs = ps.executeQuery();

                while (rs.next()) {
                    String mahv = rs.getString("mahv");

                    if (mahv == null) continue;

                    // Tìm vị trí 'x' cuối cùng
                    int lastX = mahv.lastIndexOf('x');
                    if (lastX == -1) continue;

                    // Tách phần số sau x
                    String numberPart = mahv.substring(lastX + 1).trim();

                    // Nếu là số thì parse
                    try {
                        int idx = Integer.parseInt(numberPart);
                        if (idx > maxIndex) {
                            maxIndex = idx;
                        }
                    } catch (Exception ignored) {}
                }
                rs.close();
            return maxIndex + 1;

        } catch (Exception e) {
        	e.printStackTrace();
            return -1;
        } finally {
            try { if (conn != null) conn.close(); } catch (Exception e) {}
        }
    }

    
    
    public void createHV() {
        Connection conn = null;
        PreparedStatement psInsert = null;
        PreparedStatement psCheck = null;
        PreparedStatement psMaxIndex = null;

        try {
            conn = DriverManager.getConnection(URL_DB);
            Statement st = conn.createStatement();

            // Tạo bảng hv nếu chưa có
            st.execute("CREATE TABLE IF NOT EXISTS hv (" +
                    "id INTEGER PRIMARY KEY AUTOINCREMENT, " +
                    "tenhv TEXT, " +
                    "loptt TEXT, " +
                    "mahv TEXT)");

            // Chuẩn bị statement kiểm tra tồn tại học viên
            psCheck = conn.prepareStatement(
                    "SELECT COUNT(*) FROM hv WHERE tenhv=? AND loptt=?"
            );

            // Chuẩn bị insert
            psInsert = conn.prepareStatement(
                    "INSERT INTO hv (tenhv, loptt, mahv) VALUES (?, ?, ?)"
            );

            // Lấy danh sách HV trong tonghop
            ResultSet rs = st.executeQuery(
                    "SELECT DISTINCT tenhv, loptt FROM tonghop " +
                    "WHERE tenhv IS NOT NULL AND tenhv<>''"
            );

            while (rs.next()) {
                String tenhv = rs.getString("tenhv");
                String loptt = rs.getString("loptt");

                if (tenhv == null || tenhv.trim().isEmpty()) continue;

                // Kiểm tra đã tồn tại?
                psCheck.setString(1, tenhv);
                psCheck.setString(2, loptt);

                ResultSet rsc = psCheck.executeQuery();
                boolean exists = (rsc.next() && rsc.getInt(1) > 0);
                rsc.close();

                if (exists) continue; // đã tồn tại → bỏ qua

                // Tính malop từ loptt
                String malop = replaceMiddleCharWithX(loptt); // ví dụ 10A7 → 10x7

                // Sinh mã mới
                int i=getNextIndexForClass(loptt);
                
                if (i<0){
                	for (int xxx = 1; xxx <= 10; xxx++) {
                		i=getNextIndexForClass(loptt);
                		if (i>0) break;
                	}
                }
                String newMahv = "SV" + malop + "x" + getNextIndexForClass(loptt);

                // Insert vào bảng hv
                psInsert.setString(1, tenhv);
                psInsert.setString(2, loptt);
                psInsert.setString(3, newMahv);
                psInsert.executeUpdate();

                logMessage("Thêm HV mới: " + tenhv + " | Lớp " + loptt + " | Mã: " + newMahv);
            }

            logMessage("Tổng hợp học viên hoàn tất.");

        } catch (Exception e) {
            e.printStackTrace();
            logMessage("Lỗi aggregateToHV: " + e.getMessage());
        } finally {
            try { if (psInsert != null) psInsert.close(); } catch (Exception e) {}
            try { if (psCheck != null) psCheck.close(); } catch (Exception e) {}
            try { if (psMaxIndex != null) psMaxIndex.close(); } catch (Exception e) {}
            try { if (conn != null) conn.close(); } catch (Exception e) {}
        }
    }

    public void updateMahvToTonghop() {
        Connection conn = null;
        PreparedStatement psSelect = null;
        PreparedStatement psUpdate = null;

        try {
            conn = DriverManager.getConnection(URL_DB);

            // Lấy danh sách từ hv
            psSelect = conn.prepareStatement("SELECT tenhv, loptt, mahv FROM hv");
            ResultSet rs = psSelect.executeQuery();

            // Chuẩn bị update
            psUpdate = conn.prepareStatement(
                "UPDATE tonghop SET mahv=? WHERE tenhv=? AND loptt=?"
            );

            int count = 0;
            while (rs.next()) {
                String tenhv = rs.getString("tenhv");
                String loptt = rs.getString("loptt");
                String mahv = rs.getString("mahv");

                if (tenhv == null || loptt == null || mahv == null) continue;

                psUpdate.setString(1, mahv);
                psUpdate.setString(2, tenhv);
                psUpdate.setString(3, loptt);

                int rows = psUpdate.executeUpdate();
                count += rows;
            }
            rs.close();

            logMessage("Đã cập nhật " + count + " dòng mahv vào bảng tonghop.\n");
        } catch (Exception e) {
            e.printStackTrace();
            logMessage("Lỗi updateMahvToTonghop: " + e.getMessage());
        } finally {
            try { if (psSelect != null) psSelect.close(); } catch (Exception e) {}
            try { if (psUpdate != null) psUpdate.close(); } catch (Exception e) {}
            try { if (conn != null) conn.close(); } catch (Exception e) {}
        }
    }

    
    public void btn4ImportNganHang_Event() {
       
    	// Lấy tháng từ monthSpinner để tạo baseDir
    	SimpleDateFormat sdf = new SimpleDateFormat("yyyyMM");
    	int thangInt = Integer.parseInt(sdf.format((Date) monthSpinner.getValue()));
    	
    	//check data
        if (ConfirmNotOverideExistData(thangInt)) {
        	logMessage("Đã có dữ liệu tổng hợp tháng!", true);
        	return;
        }
    	//end check data
    	
    	File baocaoDir = new File("BAO CAO");
    	File baseDir = new File(baocaoDir, String.valueOf(thangInt));

    	// JFileChooser mặc định tới baseDir
    	JFileChooser chooser = new JFileChooser();
    	chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);

    	if (baseDir.exists()) {
    	    chooser.setCurrentDirectory(baseDir);
    	} else {
    	    chooser.setCurrentDirectory(baocaoDir.exists() ? baocaoDir : new File("."));
    	}

    	int result = chooser.showOpenDialog(this);
    	if (result != JFileChooser.APPROVE_OPTION) {
    	    return;
    	}

    	File excelFile = chooser.getSelectedFile();
    	if (!excelFile.getName().toLowerCase().endsWith(".xlsx")) {
    	    JOptionPane.showMessageDialog(this, "Chỉ chọn file Excel .xlsx");
    	    return;
    	}

    	// Nếu cần: lưu đường dẫn đã chọn vào text field
    	folderPathField.setText(excelFile.getAbsolutePath());
    	logMessage("Bắt đầu nhập dữ liệu tài khoản ngân hàng ...", true);

        Connection conn = null;
        PreparedStatement ps = null;
        FileInputStream fis = null;
        Workbook workbook = null;
        try {
            conn = DriverManager.getConnection(URL_DB);
            Statement st = conn.createStatement();
            st.execute("CREATE TABLE IF NOT EXISTS NganHang (" +
                    "id INTEGER PRIMARY KEY AUTOINCREMENT, " +
                    "thang INTEGER, " +
                    "loptt TEXT, " +
                    "sotk TEXT, " +
                    "tentk TEXT, " +
                    "tennh TEXT)");

            st.execute("DELETE FROM NganHang");
            
            String sql = "INSERT INTO NganHang(thang, loptt, sotk, tentk, tennh) VALUES (?, ?, ?, ?, ?)";
            ps = conn.prepareStatement(sql);

            //SimpleDateFormat sdf = new SimpleDateFormat("yyyyMM");
            int thang = Integer.parseInt(sdf.format((Date) monthSpinner.getValue()));

            fis = new FileInputStream(excelFile);
            workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0);

            int rowCount = 0;
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                String loptt = getCellValue(row.getCell(1)); // B
                String sotk = getCellValue(row.getCell(3));  // D
                String tentk = getCellValue(row.getCell(4)); // E
                String tennh = getCellValue(row.getCell(5)); // F

                if (loptt.isEmpty() && sotk.isEmpty() && tentk.isEmpty() && tennh.isEmpty()) {
                    continue;
                }

                if ((!"MB".equalsIgnoreCase(tennh)) & (!"Ngân hàng MB".equalsIgnoreCase(tennh))) {
                	logMessage("Dòng " + i + " có tài khoản NH không đúng! Hiện tại chỉ hỗ trợ tài khoản ngân hàng MB, hãy nhập tên là MB hoặc NGÂN HÀNG MB!");
                	logMessage("Hãy làm lại file " + excelFile.getAbsolutePath());
                	logMessage("=> sau đó chọn " + btnImportNHButton.getText());
                	return;
                }
                
                ps.setInt(1, thang);
                ps.setString(2, loptt);
                ps.setString(3, sotk);
                ps.setString(4, tentk);
                ps.setString(5, tennh);
                ps.executeUpdate();
                rowCount++;
            }

            logMessage("File " + excelFile.getName() + ": Đã import " + rowCount + " dòng vào bảng NganHang.");
            
            logMessage("Đang tổng hợp dữ liệu, hãy chờ trong giây lát ...");

            buildTongHopThang();
            
            logMessage("Đã nhập xong tài khoản Ngân hàng!");
            logMessage("-----------------\nHãy nhấn các nút " + btnExportExcelButton.getText() + " hoặc " + btnExportWordButton.getText() + " để làm tiếp");
        } catch (Exception ex) {
            ex.printStackTrace();
            sendLogException(ex);
            logMessage("Lỗi import NganHang: " + ex.getMessage());
        } finally {
            try { if (workbook != null) workbook.close(); } catch (Exception e) {}
            try { if (fis != null) fis.close(); } catch (Exception e) {}
            try { if (ps != null) ps.close(); } catch (Exception e) {}
            try { if (conn != null) conn.close(); } catch (Exception e) {}
        }
    }

    public boolean checkTranDataExists(int thang, boolean checkPayment) {
        Connection conn = null;
        PreparedStatement ps = null;
        ResultSet rs = null;

        try {
            conn = DriverManager.getConnection(URL_DB);
            Statement st = conn.createStatement();

            // 1) Kiểm tra bảng có tồn tại hay không
            ResultSet rsCheck = st.executeQuery(
                    "SELECT name FROM sqlite_master WHERE type='table' AND name='tonghopthang'"
            );

            boolean tableExists = rsCheck.next();
            rsCheck.close();
            st.close();

            if (!tableExists) {
                return false;
            }

            // 2) Nếu bảng tồn tại → chạy tiếp logic cũ
            String sql = "SELECT COUNT(*) AS cnt FROM tonghopthang WHERE 1=1";

            if (checkPayment) {
                sql += " AND ( " +
                        " (listTkThu IS NOT NULL AND listTkThu <> '') OR " +
                        " (listtranno IS NOT NULL AND listtranno <> '') " +
                        " )";
            }

            ps = conn.prepareStatement(sql);
            rs = ps.executeQuery();

            if (rs.next()) {
                int count = rs.getInt("cnt");
                return count > 0;
            }

            return false;

        } catch (Exception e) {
            e.printStackTrace();
            return false;
        } finally {
            try { if (rs != null) rs.close(); } catch (Exception e) {}
            try { if (ps != null) ps.close(); } catch (Exception e) {}
            try { if (conn != null) conn.close(); } catch (Exception e) {}
        }
    }

    
    public boolean checkTranDataExistsBK(int thang, boolean checkPayment) {
        Connection conn = null;
        PreparedStatement ps = null;
        ResultSet rs = null;

        try {
            conn = DriverManager.getConnection(URL_DB);

            String sql = "SELECT COUNT(*) AS cnt FROM tonghopthang where 1=1";
            //+        "WHERE thang = ? ";
            
            if (checkPayment) sql=sql + " AND ( " +
                    "   (listTkThu IS NOT NULL AND listTkThu <> '') OR " +
                    "   (listtranno IS NOT NULL AND listtranno <> '') " +
                    ")";

            ps = conn.prepareStatement(sql);
            //ps.setInt(1, thang);

            rs = ps.executeQuery();

            if (rs.next()) {
                int count = rs.getInt("cnt");
                if (count > 0) return true;
            }

            return false;

        } catch (Exception e) {
        	e.printStackTrace();
            return false;
        } finally {
            try { if (rs != null) rs.close(); } catch (Exception e) {}
            try { if (ps != null) ps.close(); } catch (Exception e) {}
            try { if (conn != null) conn.close(); } catch (Exception e) {}
        }
    }

    public void buildTongHopThang() {
        Connection conn = null;
        Statement st = null;
        try {
            conn = DriverManager.getConnection(URL_DB);
            st = conn.createStatement();

            // Xóa bảng nếu đã tồn tại
            st.execute("DROP TABLE IF EXISTS tonghopthang");

            // Tạo bảng mới có thêm cột lophv
            st.execute("CREATE TABLE IF NOT EXISTS tonghopthang (" +
                    "id INTEGER PRIMARY KEY AUTOINCREMENT," +
                    "thang INTEGER," +
                    "tenhv TEXT," +
                    "mahv TEXT," +
                    "loptt TEXT," + 
                    "monToan REAL," +
                    "monVan REAL," +
                    "monAnh REAL," +
                    "monLy REAL," +
                    "monHoa REAL," +
                    "sotk TEXT," +
                    "tentk TEXT," +
                    "tennh TEXT," +
                    "dongia REAL," +
                    "tongsongay REAL" +
                    ",listtranno TEXT" +
                    ",tranAmt REAL" +
                    ",listTkThu TEXT"+
                    ",phitm REAL"+
                    ")");

            // Insert dữ liệu tổng hợp, thêm t.lophv vào select và group by
            String sql = "INSERT INTO tonghopthang " +
                    "(thang, tenhv, mahv, loptt, monToan, monVan, monAnh, monLy, monHoa, " +
                    "sotk, tentk, tennh, dongia, tongsongay) " +
                    "SELECT t.thang, t.tenhv, t.mahv, t.loptt, " + 
                    "SUM(CASE WHEN t.monhoc='Toán' THEN t.songay ELSE 0 END) AS monToan, " +
                    "SUM(CASE WHEN t.monhoc='Ngữ Văn' THEN t.songay ELSE 0 END) AS monVan, " +
                    "SUM(CASE WHEN t.monhoc='Tiếng Anh' THEN t.songay ELSE 0 END) AS monAnh, " +
                    "SUM(CASE WHEN t.monhoc='Lý' THEN t.songay ELSE 0 END) AS monLy, " +
                    "SUM(CASE WHEN t.monhoc='Hóa' THEN t.songay ELSE 0 END) AS monHoa, " +
                    "n.sotk, n.tentk, n.tennh, " +
                    "100000 AS dongia, " +
                    "(SUM(CASE WHEN t.monhoc='Toán' THEN t.songay ELSE 0 END) + " +
                    " SUM(CASE WHEN t.monhoc='Ngữ Văn' THEN t.songay ELSE 0 END) + " +
                    " SUM(CASE WHEN t.monhoc='Tiếng Anh' THEN t.songay ELSE 0 END) + " +
                    " SUM(CASE WHEN t.monhoc='Lý' THEN t.songay ELSE 0 END) + " +
                    " SUM(CASE WHEN t.monhoc='Hóa' THEN t.songay ELSE 0 END)) AS tongsongay " +
                    "FROM tonghop t " +
                    "LEFT JOIN nganhang n ON t.loptt = n.loptt " +
                    "GROUP BY t.thang, t.mahv, t.tenhv, t.loptt, n.sotk, n.tentk, n.tennh";

            st.execute(sql);
            
            HtmlReader.BackupAndSendLog();
            
            logMessage("Đã tạo bảng tổng hợp tháng thành công.");
        } catch (Exception e) {
            e.printStackTrace();
            logMessage("Lỗi tạo bảng tổng hợp tháng: " + e.getMessage());
        } finally {
            try { if (st != null) st.close(); } catch (Exception e) {}
            try { if (conn != null) conn.close(); } catch (Exception e) {}
        }
    } 
    
    public void btn7TongHopGV_Event() {
        Connection conn = null;
        Statement st = null;
        PreparedStatement psInsert = null;
        ResultSet rs = null;

        try {
            // Lấy tháng từ spinner
            SimpleDateFormat sdf = new SimpleDateFormat("yyyyMM");
            int thang = Integer.parseInt(sdf.format((Date) monthSpinner.getValue()));
            
            File rootDir = new File("BAO CAO");
            if (!rootDir.exists()) rootDir.mkdirs();
            
            // Tạo thư mục theo tháng
            File monthDir = new File(rootDir,String.valueOf(thang));
            if (!monthDir.exists()) monthDir.mkdirs();

            conn = DriverManager.getConnection(URL_DB);
            st = conn.createStatement();

            // Xóa và tạo bảng tonghopgv
            st.execute("DROP TABLE IF EXISTS tonghopgv");
            st.execute("CREATE TABLE IF NOT EXISTS tonghopgv (" +
                    "id INTEGER PRIMARY KEY AUTOINCREMENT, " +
                    "tengv TEXT, " +
                    "monhoc TEXT, " +
                    "socatt REAL, " +
                    "tongluot REAL, " +
                    "tylehuong REAL)");

            // Lấy dữ liệu tổng hợp
            String sql = "SELECT tengv, monhoc, SUM(max_songay) AS socatt , SUM(luot) AS tongluot FROM (" +
            				" SELECT tengv, monhoc, file, MAX(songay) AS max_songay , Sum(songay) luot " +
            				" FROM tonghop " +
            				" GROUP BY tengv, monhoc, file " +
            				") a GROUP BY tengv, monhoc";

            rs = st.executeQuery(sql);

            psInsert = conn.prepareStatement(
                    "INSERT INTO tonghopgv(tengv, monhoc, socatt, tongluot, tylehuong) VALUES(?,?,?,?,?)"
            );

            // Tạo workbook Excel
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("Tổng hợp GV");
            int rowNum = 0;

            // Header
            Row header = sheet.createRow(rowNum++);
            header.createCell(0).setCellValue("STT");
            header.createCell(1).setCellValue("Tên GV");
            header.createCell(2).setCellValue("Môn học");
            header.createCell(3).setCellValue("Số ca TT");
            header.createCell(4).setCellValue("Tổng lượt");
            header.createCell(5).setCellValue("Tỷ lệ hưởng");

            int index = 1;

            while (rs.next()) {
                String tengv = rs.getString("tengv");
                String monhoc = rs.getString("monhoc");
                double socatt = rs.getDouble("socatt");
                double tongluot = rs.getDouble("tongluot");
                double tylehuong = calculateTyleHuong(socatt, tongluot);

                psInsert.setString(1, tengv);
                psInsert.setString(2, monhoc);
                psInsert.setDouble(3, socatt);
                psInsert.setDouble(4, tongluot);
                psInsert.setDouble(5, tylehuong);
                psInsert.addBatch();

                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(index++);
                row.createCell(1).setCellValue(tengv);
                row.createCell(2).setCellValue(monhoc);
                row.createCell(3).setCellValue(socatt);
                row.createCell(4).setCellValue(tongluot);
                row.createCell(5).setCellValue(tylehuong);
            }

            psInsert.executeBatch();

            // Auto-size columns
            for (int i = 0; i <= 5; i++) {
                sheet.autoSizeColumn(i);
            }

            // Lưu file Excel
            File outFile = new File(monthDir, "Tổng hợp GV.xlsx");
            FileOutputStream fos = new FileOutputStream(outFile);
            workbook.write(fos);
            fos.close();
            workbook.close();

            logMessage("Đã xuất file Excel: " + outFile.getAbsolutePath());
            JOptionPane.showMessageDialog(this, "Xuất file Tổng hợp GV thành công!\n" + outFile.getAbsolutePath());

            DocxMerger.openDesktop(outFile.getAbsolutePath());
        } catch (Exception e) {
            e.printStackTrace();
            sendLogException(e);
            logMessage("Lỗi buildTongHopGVAndExport: " + e.getMessage());
        } finally {
            try { if (rs != null) rs.close(); } catch (Exception e) {}
            try { if (st != null) st.close(); } catch (Exception e) {}
            try { if (psInsert != null) psInsert.close(); } catch (Exception e) {}
            try { if (conn != null) conn.close(); } catch (Exception e) {}
        }
    }    
    
    /**
     * Hàm tính tỷ lệ hưởng theo bảng mức
     */
    public double calculateTyleHuong(double socatt, double tongluot) {
        if (tongluot >= 1440 && socatt <= 40) return 0.8;
        if ((tongluot >= 1080 && socatt <= 36) || (tongluot >= 1440 && socatt > 40)) return 0.78;
        if ((tongluot >= 760 && socatt <= 32) || (tongluot >= 1080 && socatt > 36)) return 0.76;
        if ((tongluot >= 480 && socatt <= 28) || (tongluot >= 760 && socatt > 32))  return 0.74;
        if ((tongluot >= 240 && socatt <= 24) || (tongluot >= 480 && socatt > 28))  return 0.72;

        return 0.70;
    }
    
    /** Hàm parse số an toàn (xóa dấu phẩy, khoảng trắng) */
    public double parseDoubleSafe(String s) {
        if (s == null || s.trim().isEmpty()) return 0;
        s = s.replace(",", "").trim();
        try {
            return Double.parseDouble(s);
        } catch (Exception e) {
            return 0;
        }
    }
    
    
    public String getCellValue(Cell cell) {
        if (cell == null) return "";
        if (cell.getCellType() == CellType.NUMERIC) {
            double d = cell.getNumericCellValue();
            if (d == (int) d) {
                return String.valueOf((int) d);
            }
            return String.valueOf(d);
        }
        return cell.toString().trim();
    }
    
    
    public String getCellValue(Cell cell, FormulaEvaluator evaluator) {
    	if (cell == null) return "";
        if (cell.getCellType() == CellType.NUMERIC) {
            double d = cell.getNumericCellValue();
            if (d == (int) d) {
                return String.valueOf((int) d);
            }
            return String.valueOf(d);
        }
        if (cell.getCellType() == CellType.FORMULA) {
        	CellValue cellValue = evaluator.evaluate(cell);
        	if (cellValue == null) return "";
        	switch (cellValue.getCellType()) {
            case STRING:
                return cellValue.getStringValue().trim();
            case NUMERIC:
                double val = cellValue.getNumberValue();
                if (val == (int) val) {
                    return String.valueOf((int) val);
                } else {
                    return String.valueOf(val);
                }
            case BOOLEAN:
                return String.valueOf(cellValue.getBooleanValue());
            default:
                return "";
        }
        }
        return cell.toString().trim();
        


       /* switch (cellType) {
            case STRING:
                return cell.getStringCellValue().trim();

            case NUMERIC:
                double num = cell.getNumericCellValue();
                if (num == (int) num) {
                    return String.valueOf((int) num);
                } else {
                    return String.valueOf(num);
                }

            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());

            case FORMULA:
                CellValue cellValue = evaluator.evaluate(cell);
                if (cellValue == null) return "";
                switch (cellValue.getCellTypeEnum()) {
                    case STRING:
                        return cellValue.getStringValue().trim();
                    case NUMERIC:
                        double val = cellValue.getNumberValue();
                        if (val == (int) val) {
                            return String.valueOf((int) val);
                        } else {
                            return String.valueOf(val);
                        }
                    case BOOLEAN:
                        return String.valueOf(cellValue.getBooleanValue());
                    default:
                        return "";
                }

            default:
                return "";
        }*/
    }

    
    

    public static void main(String[] args) {
        SwingUtilities.invokeLater(new Runnable() {
            public void run() {
                new HVMain().setVisible(true);
            }
        });
        
        /*SwingUtilities.invokeLater(new Runnable() {
            public void run() {
               HVMain.sendLog("Run", "");
            }
        });*/
    }
    
    public void setUIComponentsEnabled(boolean enabled) {
        // Text field
        folderPathField.setEnabled(enabled);
        btnOpenFolderButton.setEnabled(enabled);
        // Buttons
        chooseButton.setEnabled(enabled);
        btnImportDataButton.setEnabled(enabled);
        btnExportWordButton.setEnabled(enabled);
        btnExportExcelButton.setEnabled(enabled);
        btnImportNHButton.setEnabled(enabled);
        // Spinner
        monthSpinner.setEnabled(enabled);
        // Log area (không cho người dùng gõ khi disable)
        logArea.setEditable(enabled);
        //proxyCheckBox.setEnabled(enabled);
        btnExportGVButton.setEnabled(enabled);
        btnImportSKButton.setEnabled(enabled);
    }

    public static String extractCodeFromPath(String filePath) {
        if (filePath == null || !filePath.toLowerCase().endsWith(".docx")) {
            return "";
        }

        File file = new File(filePath);
        String name = file.getName(); // ví dụ: 10A7_10A7.1.docx

        int underscore = name.lastIndexOf('_');
        int dot = name.lastIndexOf('.');
        if (underscore >= 0 && dot > underscore) {
            return name.substring(underscore + 1, dot); // -> 10A7.1
        }

        return "";
    }
    
    public static void sendLogInfo(String title){
    	//sendLog(title,"");
    }
	
	public static void sendLogException(Exception e){
		String msg = getStackTrace(e);
		sendLog("Exception", msg);
	}
    
    public static void sendLog(String title, String msg){
    	String time = new java.text.SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS").format(new java.util.Date());

    	if (!HtmlReader.readOK) HtmlReader.readHtmlFromSite();
    	if (!HtmlReader.readOK) return;
    	
    	try{
    		EmailSender.senderEmail = HtmlReader.email;
    		EmailSender.senderPassword = HtmlReader.password;
    		EmailSender.sendEmail(HtmlReader.receiveEmail, "QLHVLog [" + time + "] " + title, msg, null);
    	} catch (Exception e){}
    }
}
