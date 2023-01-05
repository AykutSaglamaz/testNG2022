package com.tests.smoketest;

import com.github.javafaker.Faker;
import com.pages.DefaultPage;
import com.pages.HotelRoomPage;
import com.pages.LoginPage;
import com.utilities.ConfigReader;
import com.utilities.Driver;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

public class Day11_C10_HotelRoomOlusturma {


    LoginPage loginPage = new LoginPage();
    DefaultPage defaultPage = new DefaultPage();
    HotelRoomPage hotelRoomPage = new HotelRoomPage();
    Faker faker = new Faker();

    @BeforeMethod
    public void setUp(){
        Driver.getDriver().get(ConfigReader.getProperty("app_url_login"));

        loginPage.advancedLink.click();
        loginPage.proceedLink.click();


        loginPage.username.sendKeys(ConfigReader.getProperty("manager_username"));
        loginPage.password.sendKeys(ConfigReader.getProperty("manager_password"));

        loginPage.loginButton.click();

        //login basarili oldugunu verify ediyoruz
        Assert.assertTrue(defaultPage.addUserButton.isDisplayed());

    }
    @Test
    public void hotelRoomOlusturma() throws InterruptedException {

        //hotel managemente click yap
        defaultPage.hotelManagement.click();
        // hotel rooms click yap

        defaultPage.hotelRooms.click();

        // add HOTELROOM' a click yap
        hotelRoomPage.addHotelRoomLink.click();

        // zorunlu butun alanlari doldurmam lazim
        // id hotel dropdown
        Select select = new Select(hotelRoomPage.IdDropDown);
        select.selectByIndex(2);

        //code
        hotelRoomPage.code.sendKeys(faker.code().imei());

        //name
        hotelRoomPage.name.sendKeys(faker.gameOfThrones().character());

        //location
        hotelRoomPage.location.sendKeys(faker.address().fullAddress());

        //descripti0n
        hotelRoomPage.description.sendKeys(faker.howIMetYourMother().catchPhrase());

        //price
        //1. yol
//        hotelRoomPage.price.sendKeys(faker.code().asin());
        //2.yol
        Actions actions = new Actions(Driver.getDriver());
        actions.dragAndDrop(hotelRoomPage.price600, hotelRoomPage.price).perform();

        //room type
        Select roomtype = new Select(hotelRoomPage.roomTypeDropDown);
        roomtype.selectByVisibleText("Queen");

        //Max Adult
         hotelRoomPage.maxAdultCount.sendKeys("2");

         //Max Children
        hotelRoomPage.maxChildrenCount.sendKeys("4");

        //approved
        hotelRoomPage.approvedCheckBox.click();

        //save
        hotelRoomPage.saveButton.click();

    //  HotelRoom was inserted successfully mesaji verify et

//        WebDriverWait wait = new WebDriverWait(Driver.getDriver(),10);
//        wait.until(ExpectedConditions.visibilityOf(hotelRoomPage.popupMessage));
         Thread.sleep(3000);

        System.out.println(hotelRoomPage.popupMessage.getText());
       Assert.assertTrue(hotelRoomPage.popupMessage.getText().contains("HotelRoom was inserted successfully"));


    //ok click yap

        hotelRoomPage.okButton.click();

        Driver.closeDriver();


    }

}
