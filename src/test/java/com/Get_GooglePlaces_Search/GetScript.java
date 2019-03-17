package com.Get_GooglePlaces_Search;

import com.genericfunctions.GenericMethods;

import io.restassured.RestAssured;
import io.restassured.response.Response;

public class GetScript extends GenericMethods{
	
	
	public static String getPlaceDetails()
	{
		
			
		
String baseurl=getData("Get", "BaseURI");
		RestAssured.baseURI=baseurl;
		
		
		// Get all the parameters from the execl
		
		String query=getData("Get", "Query");
		String location=getData("Get", "Location");
		String radius=getData("Get", "Radius");
		String key=getData("Get", "Key");
		
		// get resourcses
		String resources=getData("Get", "Resources");
		Response response=RestAssured.given().
		
							param("query", query).
							param("location", location).
							param("radius", radius).
							param("key", key).
		
		
		
		when().
		
							get(resources).
		
		then().
		
							assertThat().statusCode(200).
		
							
		extract().response();
		return response.asString();
		
		
		
		
		
	}
	
	
	
	
	

}
