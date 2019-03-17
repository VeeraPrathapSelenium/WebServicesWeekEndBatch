package com.Reporting;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.Date;

import com.genericfunctions.GenericMethods;
import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

public class GenerateReports {
	
	public static ExtentReports extent;
	
	public static ExtentTest Test;
	
	
	public static void intializeReport()
	{
		extent=new ExtentReports(checkResultFolder()+"\\"+GenericMethods.currentclass+".html");
		
		extent.addSystemInfo("Project Name","Twitter Automation").
		addSystemInfo("Sprint","Sprint 32").
		addSystemInfo("Environment","QA");
	}
	
	
	public static void startReport()
	{
		Test=extent.startTest(GenericMethods.currentclass);
	}
	
	
	public static void logEvent(String status,String description)
	{
		switch(status.toLowerCase())
		{
		
		case "pass":
			
			Test.log(LogStatus.PASS, description);
			break;
		case "fail":
			
			Test.log(LogStatus.FAIL, description);
			break;	
			
		case "warning":
			
			Test.log(LogStatus.WARNING, description);
			break;		
		
		}
	}
	
	
	
	
	public static void endReport()
	{
		extent.flush();
		
		extent.endTest(Test);
		
		extent.close();
		
	}
	
	
	
	public static String checkResultFolder()
	{

		SimpleDateFormat sdf=new SimpleDateFormat("dd-MM-YYYY");
		
		Date d=new Date();
		
		
		System.out.println(sdf.format(d).toString());
		
		String respath=System.getProperty("user.dir")+"\\Results\\"+sdf.format(d).toString();
		
		
		File f=new File(respath);
		
		if(!f.exists() )
		{
			f.mkdirs();
		}
		
		
		
		
		return respath;
	}
	
	
	
	

}
