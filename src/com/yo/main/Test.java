package com.yo.main;

import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.htmlunit.HtmlUnitDriver;

public class Test {

	public static void main(String[] args) {
		//WebDriver wd=new FirefoxDriver();
		WebDriver driver=new HtmlUnitDriver();
		driver.get("http://172.28.136.19:8033/mlb/");
		List<WebElement> elements = (List<WebElement>) driver.findElements(By.className("aCursor"));	
		for (WebElement webElement : elements) {
			String text = webElement.getText();
			System.out.println(text);
			webElement.click();
			driver.findElement(By.id("inputWorkId")).sendKeys("w16001653");
			driver.findElement(By.id("inputPassword")).sendKeys("819513015");
			driver.findElement(By.id("loginButtom")).click();
			List<WebElement> as=driver.findElements(By.tagName("a"));
			for (WebElement webElement2 : as) {
				System.out.println(webElement2.getText());
			}
			List<WebElement> h1s =driver.findElements(By.tagName("h1"));
			for (WebElement webElement2 : h1s) {
				System.out.println(webElement2.getText().toString());
			}
			System.out.println("end");
			//不能識別中文...真是極好的,暫時不使用

			
			
		}
	}

}
