package com.javascriptexecutor;

import com.pages.LoginPage;
import com.pages.MainPage;
import com.utilities.ConfigReader;
import com.utilities.Driver;
import com.utilities.JSUtils;
import com.utilities.ReusableMethods;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.testng.Assert;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.Test;

public class Day14_C15_JSUtilTest {

    MainPage mainPage;

    @Test
    public void scrollIntoView(){
        Driver.getDriver().get(ConfigReader.getProperty("app_url"));
        mainPage = new MainPage();

        try {
            mainPage.advancedLink.click();
            mainPage.proceedLink.click();
        }catch (Exception e){
            System.out.println("Advenced linki gorunmedi");
        }
        WebElement haveAQuestion = Driver.getDriver().findElement(By.xpath("//*[.='Have a Questions?']"));

        JSUtils.scrollIntoViewJS(haveAQuestion);
        ReusableMethods.waitFor(3);

        Assert.assertEquals(haveAQuestion.getText(), "Have a Questions?" );

        Assert.assertTrue(haveAQuestion.isDisplayed());

//        Driver.closeDriver();

    }

    @Test
    public void clickByJS(){
        Driver.getDriver().get(ConfigReader.getProperty("app_url"));
        mainPage = new MainPage();
        try {
            mainPage.advancedLink.click();
            mainPage.proceedLink.click();
        }catch (Exception e){
            System.out.println("Advenced linki gorunmedi");
        }

//        WebElement loginLink =Driver.getDriver().findElement(By.linkText("Log in"));
//        JSUtils.clickElementByJS(loginLink);

        WebElement checkAvailability =Driver.getDriver().findElement(By.xpath("//input[@type='submit']"));
        JSUtils.clickElementByJS(checkAvailability);

//        Driver.closeDriver();

    }
    @Test
    public void flashJS(){
        Driver.getDriver().get(ConfigReader.getProperty("app_url"));

        mainPage = new MainPage();
        try {
            mainPage.advancedLink.click();
            mainPage.proceedLink.click();
        }catch (Exception e){
            System.out.println("Advenced linki gorunmedi");
        }

        WebElement checkAvailabilityButton =Driver.getDriver().findElement(By.xpath("//input[@type='submit']"));
        JSUtils.flash(checkAvailabilityButton);

//        Driver.closeDriver();

    }
    @Test
    public void changeColor(){
            Driver.getDriver().get(ConfigReader.getProperty("app_url"));

            mainPage = new MainPage();
            try {
                mainPage.advancedLink.click();
                mainPage.proceedLink.click();
            }catch (Exception e){
                System.out.println("Advenced linki gorunmedi");
            }

        WebElement checkAvailabilityButton =Driver.getDriver().findElement(By.xpath("//input[@type='submit']"));
            JSUtils.changeColor("red", checkAvailabilityButton);
    }

    @AfterMethod
    public void closeDriver(){
        Driver.closeDriver();
    }

}
