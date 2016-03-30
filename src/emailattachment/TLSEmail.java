/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package emailattachment;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;
 
import javax.mail.Authenticator;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
 
public class TLSEmail {
 
    /**
       Outgoing Mail (SMTP) Server
       requires TLS or SSL: smtp.gmail.com (use authentication)
       Use Authentication: Yes
       Port for TLS/STARTTLS: 587
     * @param args
     * @throws java.io.FileNotFoundException
     */
    public static void main(String[] args) throws FileNotFoundException, IOException {
        
         //connect to excel sheet
          File excel = new File("C:\\Users\\samuel.samuel-andoh\\Downloads\\TestData.xlsx");
          FileInputStream fis = new FileInputStream(excel);

                       XSSFWorkbook wb = new XSSFWorkbook(fis);
                       XSSFSheet ws = wb.getSheet("Sheet1");

                       int rowNum = ws.getLastRowNum() + 1;
                       int colNum = ws.getRow(0).getLastCellNum();
                       String[][] data = new String[(rowNum - 1)][colNum];

          for (int i=0; i <= ws.getLastRowNum(); i++){
 
          String keyword = ws.getRow(i).getCell(1).getStringCellValue();
          
         System.out.println(keyword);
        
        
        
        
        
        
        final String fromEmail = "andoh.samuel@gmail.com"; //requires valid gmail id
        final String password = "............"; // correct password for gmail id
        final String toEmail = (keyword); // can be any email id 
        
        System.out.println("TLSEmail Start");
        Properties props = new Properties();
        props.put("mail.smtp.host", "smtp.gmail.com"); //SMTP Host
        props.put("mail.smtp.port", "587"); //TLS Port
        props.put("mail.smtp.auth", "true"); //enable authentication
        props.put("mail.smtp.starttls.enable", "true"); //enable STARTTLS
        Authenticator auth;
        auth = new Authenticator() {
            //override the getPasswordAuthentication method
            @Override
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(fromEmail, password);
            }
        };
        Session session = Session.getInstance(props, auth);
         
        EmailUtil.sendAttachmentEmail(session, toEmail,"TLSEmail Testing Subject", " Attachment TLSEmail Testing Body");
         
    }
 
     
    
    
}
}