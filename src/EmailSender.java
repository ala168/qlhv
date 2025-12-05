import java.io.File;
import java.util.Properties;

import jakarta.mail.*;
import jakarta.mail.internet.*;

public class EmailSender {

    public static void main(String[] args) {

    			
    	System.setProperty("https.proxyHost", "10.53.120.1");
        System.setProperty("https.proxyPort", "8080");
        
        System.setProperty("http.proxyHost", "10.53.120.1");
        System.setProperty("http.proxyPort", "8080");
    			
        System.setProperty("socksProxyHost", "10.53.120.1");
        System.setProperty("socksProxyPort", "8080");
        
    	
	        // Ná»™i dung HTML máº«u
	        String emailBody = "test_email";
	        String fileName = "test_email";
	        // Gá»­i email
	        String toEmail = "tiepvk@outlook.com"; // TODO: chá»‰nh láº¡i
	        String subject = "Test Email tá»« Java";
	        sendEmail(toEmail, subject, emailBody);
	        sendEmail(toEmail, subject, emailBody,"E:/workspace/profile_workspace/QLHV/QLHV/data/hvtest.zip");
    }
    
    // Email vÃ  password gá»­i Ä‘i (App Password náº¿u lÃ  Gmail)
    public static String senderEmail = "tiepvkbk@gmail.com"; 
    public static String senderPassword = "kewydvycvtubhnrl";
    public static void sendEmail(String recipientEmail, String subject, String htmlBody, String attachFilePath) {
    	
    	if (HVMain.hasProxy){
    		System.setProperty("https.proxyHost", "10.53.120.1");
            System.setProperty("https.proxyPort", "8080");
            
            System.setProperty("http.proxyHost", "10.53.120.1");
            System.setProperty("http.proxyPort", "8080");
        			
            System.setProperty("socksProxyHost", "10.53.120.1");
            System.setProperty("socksProxyPort", "8080");
    	}
    	
        // Cấu hình SMTP của Gmail
        Properties props = new Properties();
        props.put("mail.smtp.auth", "true");
        props.put("mail.smtp.starttls.enable", "true");
        props.put("mail.smtp.host", "smtp.gmail.com");
        props.put("mail.smtp.port", "587");

        // Xác thực người gửi
        Session session = Session.getInstance(props, new Authenticator() {
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(senderEmail, senderPassword);
            }
        });

        try {
            // Tạo email
            Message message = new MimeMessage(session);
            message.setFrom(new InternetAddress(senderEmail));
            message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(recipientEmail));
            message.setSubject(subject);

            // Tạo phần body (HTML)
            MimeBodyPart bodyPart = new MimeBodyPart();
            bodyPart.setContent(htmlBody, "text/html; charset=utf-8");

            // Multipart chứa body và file
            Multipart multipart = new MimeMultipart();
            multipart.addBodyPart(bodyPart);

            // Nếu có file đính kèm
            if (attachFilePath != null && !attachFilePath.isEmpty()) {
                MimeBodyPart attachPart = new MimeBodyPart();
                attachPart.attachFile(new File(attachFilePath));
                multipart.addBodyPart(attachPart);
            }

            // Gắn multipart vào message
            message.setContent(multipart);

            // Gửi email
            Transport.send(message);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    
    public static void sendEmail(String recipientEmail, String subject, String htmlBody) {
        // Cáº¥u hÃ¬nh SMTP cá»§a Gmail
        Properties props = new Properties();
        props.put("mail.smtp.auth", "true");
        props.put("mail.smtp.starttls.enable", "true");
        props.put("mail.smtp.host", "smtp.gmail.com");
        props.put("mail.smtp.port", "587");

        // XÃ¡c thá»±c ngÆ°á»�i gá»­i
        Session session = Session.getInstance(props, new Authenticator() {
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(senderEmail, senderPassword);
            }
        });

        try {
            // Táº¡o email
            Message message = new MimeMessage(session);
            message.setFrom(new InternetAddress(senderEmail));
            message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(recipientEmail));
            message.setSubject(subject);
            message.setContent(htmlBody, "text/html; charset=utf-8");

            // Gá»­i email
            Transport.send(message);
            System.out.println("âœ… Email Ä‘Ã£ Ä‘Æ°á»£c gá»­i tá»›i: " + recipientEmail);

        } catch (MessagingException e) {
            System.err.println("â�Œ Lá»—i khi gá»­i email: " + e.getMessage());
        }
    }

}
