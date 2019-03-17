package com.TestCases;

import org.testng.annotations.Test;

import com.Get_GooglePlaces_Search.GetScript;
import com.genericfunctions.GenericMethods;

import io.restassured.RestAssured;
import io.restassured.path.json.JsonPath;
import io.restassured.response.Response;

public class Tc01_GetPlaceDetails_GogleMaps extends GenericMethods{

	
	@Test
	
	public static void getPlaceDetails()
	{
		
		GetScript getscript=new GetScript();
		String response=GetScript.getPlaceDetails();
		
		
		JsonPath jsonpath=getRawJson(response);
System.out.println(jsonpath.getString("results.size()"));
		
		// print all the places 
		
		
		for(int i=0;i<=Integer.parseInt(jsonpath.getString("results.size()"))-1;i++)
		{
			System.out.println(jsonpath.getString("results["+i+"].formatted_address"));
			
			System.out.println();
			
			System.out.println("==================================================");
			
			if(! jsonpath.getString("results["+i+"].formatted_address").isEmpty())
			{
				logEvent("pass", "Result Set "+(i+1)+" : "+jsonpath.getString("results["+i+"].formatted_address"));
			}else
			{
				logEvent("fail", "Result Set "+(i+1)+" : is Empty");
			}
			
			
			
			
		}




		
	}
	
	
	
	
	
	
}
