package com.crackjob.fellowcoders.SendingMailFromExcell;
import java.util.Properties;

import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

public class SendingMail {




    
    public static boolean sendMail(String toMailId,String name, int stu_slno, int courseRowCount ,
	    	int studentRowCount, int courseColumnCount ,int	 studentColumnCount   )
    {
    	
    	
    	FetchData fetchDataObj1 = new FetchData("C:\\Users\\SAYAK\\Desktop\\STUDENTS& COURSE DETAILS FOR CRACKJOB.xlsx");
        // Recipient's email ID needs to be mentioned.
       
        boolean mailSentSuccessfully = false;
            String to = toMailId;
            

            // Sender's email ID needs to be mentioned
            final String from = "sayaknandy11@gmail.com";

            // Assuming you are sending email from through gmails smtp
            String host = "smtp.gmail.com";

            // Get system properties
            Properties properties = new Properties();    

            // Setup mail server
            properties.put("mail.smtp.host", host);
            properties.put("mail.smtp.port", "587");
           
            properties.put("mail.smtp.socketFactory.port", "465");    
            properties.put("mail.smtp.socketFactory.class",    
                      "javax.net.ssl.SSLSocketFactory");    
            properties.put("mail.smtp.auth", "true");    
          //  properties.put("mail.smtp.port", "465");
          //  properties.put("mail.smtp.ssl.enable", "true");
           // properties.put("mail.smtp.auth", "true");

           
           
           
           
           
           
           
            // Get the Session object.// and pass username and password
           // Session session = Session.getDefaultInstance(properties);

            Session session = Session.getDefaultInstance(properties,    
                    new javax.mail.Authenticator() {    
                    protected javax.mail.PasswordAuthentication getPasswordAuthentication() {    
                    return new javax.mail.PasswordAuthentication(from,"siemenskasba");  
                    }    
                   });    
            // Used to debug SMTP issues
            session.setDebug(true);

            try {
            	// Create a default MimeMessage object.
            	MimeMessage message = new MimeMessage(session);

            	// Set From: header field of the header.
            	message.setFrom(new InternetAddress(from));

            	// Set To: header field of the header.
            	message.addRecipient(Message.RecipientType.TO, new InternetAddress(to));

            	// Set Subject: header field
            	message.setSubject("Student Course Enrollment");

            	MimeBodyPart messageBodyPart = new MimeBodyPart();
            	MimeMultipart multipart = new MimeMultipart();
            	messageBodyPart = new MimeBodyPart();
            	StringBuffer messagebody = new StringBuffer();
            	messagebody.append("Hi "+ name+",                         \n           <br>      ");
            	messagebody.append("<html>");
            	messagebody.append("\n\n\n");
            	messagebody.append("We are delighted to inform you that you are successfully enrolled with us for the Crack Job Batch 1 (CJB1).      Please find your details below:                  <br>              \r\n\n\n");
            	messagebody.append("\n\n\n");
            	
            	//messagebody.append("<table>");

            	//messagebody.append(messageTable);
            	//messagebody.append("\n");

            	//messagebody.append("</table>");

            	//messagebody.append("</html>");
            	//Setting id of student
            	String id_no;
            	if( stu_slno<=10)
            	id_no="CJ00"+ (stu_slno-1);
            	else if( stu_slno>10 &&  stu_slno<=100)
            	id_no="CJ0"+ (stu_slno-1);
            	else
            		id_no="CJ"+ (stu_slno-1);
            	fetchDataObj1.setCellData("STUDENT_DETAILS", "STUDENT_ID",stu_slno,id_no);
            	messagebody.append("Full Name:"+name+" "+fetchDataObj1.getCellData("STUDENT_DETAILS","Last_Name" , stu_slno)+"  <br>                                      \n\n");
            	messagebody.append("Student ID NUMBER:  "+ id_no+"       <br>                                    \n\n");
            	messagebody.append("\n\n\n");
            	int rowNum = fetchDataObj1.getCellRowNum("COURSE_DETAILS","COURSE_CODE" , fetchDataObj1.getCellData("STUDENT_DETAILS","COURSE_CODE" , stu_slno));
        		
    			//for (int m =0 ; m <courseColumnCount; m++) {
        			
        		//	fetchDataObj1.getCellData("COURSE_DETAILS",m , rowNum);
        		
               
            	
            	messagebody.append("COURSE START DATE:        "+fetchDataObj1.getCellData("COURSE_DETAILS","START DATE" , rowNum)+"   <br>             \n");
            	messagebody.append("\n");
            	messagebody.append("\n");
            	messagebody.append("\n");
            	messagebody.append("\n");
            	messagebody.append("COURSE END DATE:    "+fetchDataObj1.getCellData("COURSE_DETAILS","END DATE" , rowNum)+"            <br>                 \n\n");
            	messagebody.append("\n");
            	messagebody.append("\n");
            	messagebody.append("\n");
            	messagebody.append("\n");
            	messagebody.append("COURSE DURATION:     "+fetchDataObj1.getCellDataNumber("COURSE_DETAILS","DURATION" , rowNum)+"  months         <br>           \n\n");
            	messagebody.append("\n");
            	messagebody.append("\n");
            	messagebody.append("\n");
            	messagebody.append("\n");
            	messagebody.append("\n");
            	messagebody.append("COURSE PROVIDE:      "+fetchDataObj1.getCellData("COURSE_DETAILS","COURSE_DETAILS" , rowNum)+"       <br>                   \n\n");
            	messagebody.append("\n");
            	//ref_code setting 
            	String Ref_code= fetchDataObj1.getCellData("STUDENT_DETAILS","Last_Name" , stu_slno)+fetchDataObj1.getCellData("STUDENT_DETAILS","NAME" , stu_slno).charAt(0)+"@"+fetchDataObj1.getCellDataNumber("STUDENT_DETAILS","CONTACT _NUMBER" , stu_slno);
            	messagebody.append("\n");
            	fetchDataObj1.setCellData("STUDENT_DETAILS", "REF_CODE",stu_slno,Ref_code);
            	messagebody.append("\n");
            	messagebody.append("\n");
            	messagebody.append("\n");
            	messagebody.append("\n");
            	messagebody.append("Referral Code :     "+Ref_code+"          <br>                                                 \n\n");
            	messagebody.append("\n");
            	messagebody.append("\n");
            	messagebody.append("\n");
            	messagebody.append("Thank you for joining us and trusting our services. Please reply back for any queries or changes in your details.\r                    \n");
            	messagebody.append("\n");
            	messagebody.append("\n");
            	messagebody.append("\n");
            	messagebody.append("\n");
            	messagebody.append("\n");
            
            	messagebody.append("  <br>     Regards,                     <br>         \r\n Crack Job Team     <br>              \r\n                          Digital Education Foundation\r  <br>                     \n       Contact:8910274229 / 8240900937 <br>       \n");
            	messagebody.append("\n");
            	
            	

            	messagebody.append("</html>");

            	messageBodyPart.setContent(messagebody.toString(), "text/html");
            	multipart.addBodyPart(messageBodyPart);
            	// Now set the actual message
            	message.setContent(multipart);

            	System.out.println("sending...");
            	messagebody.append("\n");
            	// Send message
            	Transport.send(message);
                System.out.println("Sent message successfully....");
                mailSentSuccessfully = true;
            } catch (MessagingException mex) {
                mex.printStackTrace();
                mailSentSuccessfully = false;

            }
    return mailSentSuccessfully;

        }



}










