package com.postscripts;

import org.testng.annotations.Listeners;
import org.testng.annotations.Test;

import com.genericfunctions.GenericMethods;
import com.readexcel.ReadExcel;

import io.restassured.RestAssured;
import io.restassured.path.json.JsonPath;
import io.restassured.response.Response;

import static io.restassured.RestAssured.given;


public class Tc_01_GetStatus_Timeline  extends GenericMethods{
	
	
	@Test
	public static void GetStatusTimeLine()
	{	
		
		String resources=getData("Get","Resources");
		String parameters=getData("Get","Parameters");
		
		String timeline_Response=getRequest(resources,parameters);
		
		JsonPath path=getRawJson(timeline_Response);
	
		System.out.println(path.get("text").toString());
		
		
		if(path.get("text").toString().length()>0)
		{
			logEvent("pass","User is getting correct responses."+path.get("text").toString());
			
		}else
		{
			logEvent("fail","User is not getting correct responses");
		}
		
		 
		
	}
	

}
