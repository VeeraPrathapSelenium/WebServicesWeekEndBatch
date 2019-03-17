package com.genericfunctions;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;

import com.readexcel.ReadExcel;

import io.restassured.RestAssured;
import io.restassured.path.json.JsonPath;
import io.restassured.response.Response;

import static io.restassured.RestAssured.given;

public class GenericMethods extends ReadExcel{
//===================================================Variables Declaration==============================================================================
	
public static Properties property ;  

public static String currentclass;
	
	
	
//====================================================Methods Declaration===========================================================================================	
	/*
	 * Methodname : getRawJson(String jsonString,String fieldname)
	 * 
	 * Purpose : This method help to fetch the json object value 
	 * 
	 * Input Parameters : User must send json String and fieldname
	 * 
	 * Output Parameter : String
	 * 
	 * Designer : ****
	 * 
	 * Reviewer : *****
	 * 
	 * Design date : 12 oct 2018
	 * 
	 * Review date : ****
	 * 
	 * Status : Inprogress or to be review
 	 */
	public static JsonPath getRawJson(String jsonString)
	{  
		// initialise string value as null
		
		JsonPath jp = null;
		try
		{
			// creating JsonPath object 
			
			jp=new JsonPath(jsonString);
			
			
			System.out.println("Test");
		}
	    catch(Exception e)
	    {
	    	System.out.println(e.getMessage());
	    	
	    }
	return jp;
	}
	
	/*
	 * Methodname : loadPropertiesFile
	 * 
	 * Purpose : This method help to fetch the Environment properties
	 * 
	 * Input Parameters : NA
	 * 
	 * Output Parameter : NA
	 * 
	 * Designer : ****
	 * 
	 * Reviewer : *****
	 * 
	 * Design date : 12 oct 2018
	 * 
	 * Review date : ****
	 * 
	 * Status : Inprogress or to be review
 	 */
	
  public static void loadPropertiesFile() 
  
  { File propertiesfile = new File("Environment.properties");
    // Verify file existance
    if (propertiesfile.exists()) 
    {
    	FileInputStream fis;
		try {
			
			fis = new FileInputStream(propertiesfile);
			property = new Properties();
	    	property.load(fis);
	    	
		} catch (Exception e) {
			
			e.printStackTrace();
		}
    	
    	
    }
	
}

  /*
	 * Methodname : GetRequest
	 * 
	 * Purpose : This method help to get the request from the server
	 * 
	 * Input Parameters : NA
	 * 
	 * Output Parameter : String
	 * 
	 * Designer : ****
	 * 
	 * Reviewer : *****
	 * 
	 * Design date : 12 oct 2018
	 * 
	 * Review date : ****
	 * 
	 * Status : Inprogress or to be review
	 */
	
  public static String getRequest(String resources,String parameters)
  {
	  String crntresponse="";
	  try
	  {
		  RestAssured.baseURI=property.getProperty("BaseUrl");
		  
		  
		  String consumer=property.getProperty("Consumer");
		  String consumerKey=property.getProperty("ConsumerSecret");
		  String token=property.getProperty("Token");
		  String tokenKey=property.getProperty("TokenSecret");
		  
		 Response res= given().
				 auth().oauth(consumer, consumerKey, token, tokenKey).
		  queryParam("count",parameters).
		  when().
		  get(resources).then().and().extract().response();
		 crntresponse=res.asString() ;
		 
		 
	  }catch(Exception e)
	  {
		  System.out.println(e.getMessage());
	  }
	return crntresponse;
  }
  
  
  
  

}

