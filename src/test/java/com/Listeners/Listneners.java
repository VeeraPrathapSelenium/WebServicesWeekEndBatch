package com.Listeners;

import java.lang.reflect.AccessibleObject;
import java.lang.reflect.Method;


import org.omg.CORBA.portable.InvokeHandler;
import org.testng.IInvokedMethod;
import org.testng.IInvokedMethodListener;
import org.testng.IMethodInstance;
import org.testng.ITestContext;
import org.testng.ITestListener;
import org.testng.ITestNGMethod;
import org.testng.ITestResult;

import com.Reporting.GenerateReports;
import com.genericfunctions.GenericMethods;
import com.readexcel.ReadExcel;

public class Listneners implements ITestListener,IInvokedMethodListener{

	@Override
	public void onTestStart(ITestResult result) {
		
		System.out.println("on test start");
		
		
	}

	@Override
	public void onTestSuccess(ITestResult result) {
	
		
	}

	@Override
	public void onTestFailure(ITestResult result) {
	
		
	}

	@Override
	public void onTestSkipped(ITestResult result) {
		
		
	}

	@Override
	public void onTestFailedButWithinSuccessPercentage(ITestResult result) {
		
		
	}

	@Override
	public void onStart(ITestContext context) {
		GenericMethods.loadPropertiesFile();
		ReadExcel.loadExcelFile();
	}

	@Override
	public void onFinish(ITestContext context) {
		System.out.println("on finish");
		
	}

	@Override
	public void beforeInvocation(IInvokedMethod method, ITestResult testResult) {
		
		System.out.println("Before method");
		
		String crntclassname=testResult.getInstanceName().toString();
		/*System.out.println(crntclassname);
		System.out.println(crntclassname.substring(crntclassname.lastIndexOf(".")+1));*/
		GenericMethods.currentclass=crntclassname.substring(crntclassname.lastIndexOf(".")+1);
		
		GenerateReports.intializeReport();
		GenerateReports.startReport();
		
	}

	@Override
	public void afterInvocation(IInvokedMethod method, ITestResult testResult) {
		
		System.out.println("After method");
		GenerateReports.endReport();
	}


	

}
