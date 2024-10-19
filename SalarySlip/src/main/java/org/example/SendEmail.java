package org.example;

import javax.mail.*;
import javax.mail.internet.*;
import javax.activation.*;
import java.util.Properties;

public class SendEmail {
    public void sendEmailWithAttachment(String recipientEmail, String pdfFilePath) throws Exception {
        String from = "your-email@gmail.com";
        String host = "smtp.gmail.com";

        Properties properties = System.getProperties();
        properties.setProperty("mail.smtp.host", host);
        properties.setProperty("mail.smtp.port", "465");
        properties.setProperty("mail.smtp.auth", "true");
        properties.setProperty("mail.smtp.starttls.enable", "true");

        Session session = Session.getDefaultInstance(properties, new Authenticator() {
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication("your-email@gmail.com", "your-password");
            }
        });

        try {
            MimeMessage message = new MimeMessage(session);
            message.setFrom(new InternetAddress(from));
            message.addRecipient(Message.RecipientType.TO, new InternetAddress(recipientEmail));
            message.setSubject("Generated PDF Document");

            // Create the message part
            BodyPart messageBodyPart = new MimeBodyPart();
            messageBodyPart.setText("Please find the attached PDF document.");

            // Create a multipart message
            Multipart multipart = new MimeMultipart();
            multipart.addBodyPart(messageBodyPart);

            // Add the PDF attachment
            messageBodyPart = new MimeBodyPart();
            DataSource source = new FileDataSource(pdfFilePath);
            messageBodyPart.setDataHandler(new DataHandler(source));
            messageBodyPart.setFileName(pdfFilePath);
            multipart.addBodyPart(messageBodyPart);

            // Send the complete message parts
            message.setContent(multipart);

            Transport.send(message);
            System.out.println("Email sent successfully with the PDF attachment.");
        } catch (MessagingException mex) {
            mex.printStackTrace();
        }
    }
}
