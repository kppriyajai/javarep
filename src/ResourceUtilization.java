package com.resource;

import java.util.ArrayList;
import java.util.Properties;

import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.SendFailedException;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.AddressException;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;

//import com.resource.ExcelInput;

public class ResourceUtilization {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
//ExcelInput excel=new ExcelInput("C:\\EGIL37930IPSF201516\\New\\resourceutilization.xls");
//System.out.println("utilization slab column : "+excel.GetCell("Utilization Slab"));
		ArrayList<String> columndata=null;
		ArrayList<String> emailaddress=null;
		String path =args[0];
		ArrayList<String> columndatafinal=new ArrayList<String>();
		ExtractExcel extract=new ExtractExcel();
		columndata=extract.extractExcelContentByColumnIndex(21,path);
		emailaddress=extract.extractExcelContentByColumnIndex(40,path);
		for(int i=0;i<columndata.size();i++)
		{
			if(!(columndata.get(i).equals(">80%")))
					{
				//columndatafinal.add(columndata.get(i));
				//System.out.println("inside");
				columndatafinal.add(emailaddress.get(i));
			
					}
		}
		for(int j=0;j<columndatafinal.size();j++)
		{
			System.out.println(columndatafinal.get(j));
		}
		String from = "ResourceUtilization@ericsson.com";
		String host2 = "smtp.internal.ericsson.com";
		Properties properties = System.getProperties();
		//properties.clear();
		properties.setProperty("mail.smtp.host", host2);
		
		/*properties.setProperty("mail.smtp.socketFactory.class", SSL_FACTORY);
		properties.setProperty("mail.smtp.socketFactory.fallback", "false");
		properties.setProperty("mail.smtp.port", "465");
		properties.setProperty("mail.smtp.socketFactory.port", "465");
		properties.put("mail.smtp.auth", "true");
		properties.put("mail.debug", "true");
		properties.put("mail.store.protocol", "pop3");
	     props.put("mail.transport.protocol", "smtp");
	     final String username = "xxxx@gmail.com";//
	     final String password = "0000000";*/
	     
		Session session = Session.getDefaultInstance(properties);
		MimeMessage message = new MimeMessage(session);
		InternetAddress[] address = new InternetAddress[columndatafinal.size()];
	    for (int i = 0; i < columndatafinal.size(); i++) {
	        try {
				address[i] = new InternetAddress(columndatafinal.get(i));
			} catch (AddressException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
	    }
	    try {
				message.setRecipients(Message.RecipientType.TO, address);
				message.setFrom(new InternetAddress(from));
			  message.setSubject("Resource Utilization Defaulter");
			// Now set the actual message
		         message.setText("Resource Utilization is below 80%. Please Check with your LM");
		      // Send message
		         Transport.send(message);
		         System.out.println("Sent message successfully....");System.out.println("Sent message successfully....");
		} catch (SendFailedException e) {
			// TODO Auto-generated catch block
			//session.getProperties().clear();
			System.out.println("No Emails should be triggerred as all resources have utilization morethan 80%");
			//e.printStackTrace();
			
		}
	    catch (Exception e) {
			// TODO Auto-generated catch block
			//session.getProperties().clear();
			System.out.println("No Emails should be triggerred as all resources have utilization morethan 80%");
			//e.printStackTrace();
			
		}

	}

}
