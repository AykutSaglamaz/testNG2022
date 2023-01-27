package com.tests.smoketest;

import com.pages.DefaultPage;
import com.pages.LoginPage;
import com.pages.MainPage;
import com.utilities.ConfigReader;
import com.utilities.Driver;
import org.testng.Assert;
import org.testng.annotations.Test;

public class Day10_C8_PositiveLoginTest {

    @Test
    public void positiveLoginTest(){
        Driver.getDriver().get(ConfigReader.getProperty("app_url_login"));

        /*
        1-page object olustur ==> mainPage ve LoginPage
        2. test class olustur
         page object olustur
            - mainPage objectleri olustur
            - loginPage objectleri olustur
        */
        MainPage mainPage = new MainPage();
    try {
    mainPage.advancedLink.click();
    mainPage.proceedLink.click();

    mainPage.mainPageLoginLink.click();
    }catch (Exception e){

    }


        // bu nokta loginPage teyiz
       // loginPage object olustur

        LoginPage loginPage = new LoginPage();

//        loginPage.username.sendKeys("manager");
//        loginPage.password.sendKeys("Manager1!");

        loginPage.username.sendKeys(ConfigReader.getProperty("manager_username"));
        loginPage.password.sendKeys(ConfigReader.getProperty("manager_password"));

        loginPage.loginButton.click();


        //Adduser buttonunu kullarakn defaul page giris yaptigimizi verify edin
        //Bu noktada defaultPage clasindayiz

        DefaultPage defaultPage = new DefaultPage();

        boolean logedIn = defaultPage.addUserButton.isDisplayed();
        Assert.assertTrue(logedIn);

        Driver.closeDriver();
    }

}
