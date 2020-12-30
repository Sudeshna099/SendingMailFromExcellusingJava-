package com.crackjob.fellowcoders.SendingMailFromExcell.app;
import com.crackjob.fellowcoders.*;
import java.util.Properties;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Calendar;

import org.apache.poi.hssf.usermodel.*;


import javax.mail.*;

import com.crackjob.fellowcoders.SendingMailFromExcell.FetchData;
import com.crackjob.fellowcoders.SendingMailFromExcell.SendingMail;

public class AppMain {
    public static void main(String[] args){
    	
    	
   
    	
    	StringBuffer  sb = new StringBuffer();
    	
    	SendingMail send =new SendingMail();
    	FetchData fetchDataObj = new FetchData("D:\\java workspace\\SendingMailFromExcell\\STUDENTS& COURSE DETAILS FOR CRACKJOB.xlsx");
    	int courseRowCount = 0;
    	int studentRowCount = 0 ;
     	int courseColumnCount = 0;
    	int studentColumnCount = 0 ;
    	if(fetchDataObj.isSheetExist("STUDENT_DETAILS")) {
    		studentRowCount = fetchDataObj.getRowCount("STUDENT_DETAILS");
    		studentColumnCount = fetchDataObj.getColumnCount("STUDENT_DETAILS");
    		

    	}

    	if(fetchDataObj.isSheetExist("COURSE_DETAILS")) {
    		courseRowCount = fetchDataObj.getRowCount("COURSE_DETAILS");
    		courseColumnCount = fetchDataObj.getColumnCount("COURSE_DETAILS");

     	}
    	
    	if(fetchDataObj.isSheetExist("COURSE_DETAILS")) {
    		courseRowCount = fetchDataObj.getRowCount("COURSE_DETAILS");
    		courseColumnCount = fetchDataObj.getColumnCount("COURSE_DETAILS");

     	}
        StringBuffer header = new StringBuffer();
    	
    	header.append("<tr>");

         for (int j = 0 ; j < studentColumnCount ; j++) {
        		
        	 header.append("<td>");
        	 header.append(fetchDataObj.getCellData("STUDENT_DETAILS",j,1));
    			if (studentColumnCount != j) {
        			sb.append("</td>");

    			}
    
    		
    		}
         
         for (int m =0 ; m <courseColumnCount; m++) {
        	 header.append("<td>");
        	 header.append(fetchDataObj.getCellData("COURSE_DETAILS",m , 1));
        	 header.append("</td>");

			}
     	header.append("</tr>");

		

    	for (int i = 2 ; i <= studentRowCount ; i++) {
    		if ("Yes".equalsIgnoreCase(fetchDataObj.getCellData("STUDENT_DETAILS","MAIL_SENT_FLAG" , i))) {
    			continue;
    		}
    	
    		String toMailId = fetchDataObj.getCellData("STUDENT_DETAILS","EMAIL" , i);
    		String name = fetchDataObj.getCellData("STUDENT_DETAILS","NAME" , i);

    		int rowNum = fetchDataObj.getCellRowNum("COURSE_DETAILS","COURSE_CODE" , fetchDataObj.getCellData("STUDENT_DETAILS","COURSE_CODE" , i));
    		if (i == 1){
        		rowNum ++;
        		rowNum ++;
    		}
			
		
			if (i != 1) {
				 
				boolean mailSendSuccessfully = send.sendMail(toMailId,name, i, courseRowCount ,studentRowCount,courseColumnCount ,studentColumnCount );
				    if (mailSendSuccessfully) {
			    		boolean flagUpdated = fetchDataObj.setCellData("STUDENT_DETAILS", "MAIL_SENT_FLAG", i, "Yes");
				    }
				    sb = new StringBuffer();
			}
		  


    	}
	

	}
}