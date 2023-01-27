package com.dataprovider;

import com.pages.LoginPage;
import com.utilities.ConfigReader;
import com.utilities.Driver;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class Day14_C17_DataProvider2 {
     /*
     Manager kullanici bilgileri ile login fonksiyonunu test et
     getData method'unda 2D array donduren bir object olustur
     */

    @DataProvider (parallel = true)
    public Object [][] getData(){

        Object [][]  managerProfili = { {"manager",   "Manager1!"},{"manager5", "Manager5!"},{"manager12", "Manager12!"}};
        return managerProfili;
    }

    @DataProvider
    public Object [][] getDatam(){

        Object [][]  managerProfili = { {"manager",   "Manager1!"},{"manager5", "Manager5!"},{"manager12", "Manager12!"}};
        return managerProfili;
    }

    LoginPage loginPage;

    public void setUp(){
        loginPage = new LoginPage();
        Driver.getDriver().get(ConfigReader.getProperty("app_url_login"));
        try {
            loginPage.advancedLink.click();
            loginPage.proceedLink.click();
        }catch (Exception e){
            System.out.println("Advenced linki gorunmedi");
        }

    }

    @Test(dataProvider = "getData")
    public void managerLoginTest(String KullaniciAdi, String sifre){
        setUp();
        loginPage.username.sendKeys(KullaniciAdi);
        loginPage.password.sendKeys(sifre);
        loginPage.loginButton.click();

    }
    @AfterMethod
    public void tearDown(){
        Driver.closeDriver();
    }
}
