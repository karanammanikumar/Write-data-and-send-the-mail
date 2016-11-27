import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.*;
import javax.mail.internet.*;

public class SendFromYahoo
{
public static void send()
{    
    // Sender's email ID needs to be mentioned
     String from = "mani_jsd@yahoo.com";
     String pass ="himabindu@25";
    // Recipient's email ID needs to be mentioned.
   String to = "mani_jsd@yahoo.com";
   String host = "smtp.mail.yahoo.com";

   // Get system properties
   Properties properties = System.getProperties();
   // Setup mail server
   properties.put("mail.smtp.starttls.enable", "true");
   properties.put("mail.smtp.host", host);
   properties.put("mail.smtp.user", from);
   properties.put("mail.smtp.password", pass);
   properties.put("mail.smtp.port", "587");
   properties.put("mail.smtp.auth", "true");

   // Get the default Session object.
   Session session = Session.getDefaultInstance(properties);

   try{
      // Create a default MimeMessage object.
      MimeMessage message = new MimeMessage(session);

      // Set From: header field of the header.
      message.setFrom(new InternetAddress(from));

      // Set To: header field of the header.
      message.addRecipient(Message.RecipientType.TO,
                               new InternetAddress(to));

      // Set Subject: header field
      message.setSubject("Sending Excel file");

      //3) create MimeBodyPart object and set your message text     
      BodyPart messageBodyPart1 = new MimeBodyPart();  
      messageBodyPart1.setText("This is message body");  
        
      //4) create new MimeBodyPart object and set DataHandler object to this object      
      MimeBodyPart messageBodyPart2 = new MimeBodyPart();  
    
      String filename = "D:/exceldata.xlsx";//change accordingly  
      DataSource source = new FileDataSource(filename);  
      messageBodyPart2.setDataHandler(new DataHandler(source));  
      messageBodyPart2.setFileName(filename);  
       
       
      //5) create Multipart object and add MimeBodyPart objects to this object      
      Multipart multipart = new MimeMultipart();  
      multipart.addBodyPart(messageBodyPart1);  
      multipart.addBodyPart(messageBodyPart2);  
    
      //6) set the multiplart object to the message object  
      message.setContent(multipart );  
       

      // Send message
      Transport transport = session.getTransport("smtp");
      transport.connect(host, from, pass);
      transport.sendMessage(message, message.getAllRecipients());
      transport.close();
      System.out.println("Sent message successfully....");
   }catch (MessagingException mex) {
      mex.printStackTrace();
   }
}
}