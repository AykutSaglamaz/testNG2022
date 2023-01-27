package com.tests;

import com.pages.DefaultPage;
import com.pages.HotelRoomPage;
import com.pages.LoginPage;
import com.utilities.ConfigReader;
import com.utilities.Driver;
import com.utilities.ExcelUtil;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.testng.Assert;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

import java.util.List;

public class Day12_C11_WebTables {

//    When kullanici application'da HotelRoom sayfasina gider
    // Log in https://www.carettahotel.com/
    // Hotel Management'e Click yapar
    // Hotel Rooms'a Click yapar
//    test method: entireTable() ve tum header'lari yazdir
//    test method: printRows() ve tum row'lari ile 4 row'daki elementi yazdir
//    test method: printCells() ve table body'dek' cells'lerin toplam sayilarini yazdir ve tum cell'leri yazdir
//    test method: printColumns() ve columns'larin toplam sayisini yazdir ve 5. column yazdir
//    test method: printData(int row, int column); Bu method veriline cell'i yazdirmali
//    ornek : printData(2,3); 2. row ve 3 column yazdirmali

    LoginPage loginPage = new LoginPage();
    DefaultPage defaultPage = new DefaultPage();
    HotelRoomPage hotelRoomPage = new HotelRoomPage();


    @BeforeMethod
    public void setUp(){
        Driver.getDriver().get(ConfigReader.getProperty("app_url_login"));

        // private isterse
        loginPage.advancedLink.click();
        loginPage.proceedLink.click();

        loginPage.username.sendKeys(ConfigReader.getProperty("manager_username"));
        loginPage.password.sendKeys(ConfigReader.getProperty("manager_password"));
        loginPage.loginButton.click();

        //login basarili oldugunu assert edelim
        Assert.assertTrue(defaultPage.addUserButton.isDisplayed());

        // Hotel Management'e Click yapar
        defaultPage.hotelManagement.click();
        // Hotel Rooms'a Click yapar
        defaultPage.hotelRooms.click();
    }
    //    test method: entireTable() ve tum header'lari yazdir
    @Test
    public void entireTable(){
        //setUp();
        System.out.println("Tum Table");
        System.out.println("Table Body");

        WebElement tableBody =Driver.getDriver().findElement(By.xpath("//table//tbody"));
        System.out.println(tableBody.getText());

//        ve tum header'lari yazdir
           List<WebElement> tumHeader = Driver.getDriver().findElements(By.xpath("//th"));
           for (WebElement herHeader : tumHeader){
               System.out.print(herHeader.getText()+ "  ");
           }
    }

//    test method: printRows() ve tum row'larin sayisi ile 4. row'daki elementi yazdir
    @Test (groups = "test-group-1")
    public void printRows(){

        System.out.println("Print Rows");
        List<WebElement> tumRows = Driver.getDriver().findElements(By.xpath("//tbody//tr"));

        // tum row'larin sayisi
        int satirNumarasi=1;
        for (WebElement herRow : tumRows){
            System.out.println("Satir " + satirNumarasi + " => " + herRow.getText());
            satirNumarasi++;
        }

        // 4. row'daki elementi yazdir
        WebElement row4 =Driver.getDriver().findElement(By.xpath("//tbody//tr[4]"));
        System.out.println("Row 4 => "+ row4.getText());
        Driver.closeDriver();

    }
//    test method: printCells() ve table body'dek' cells'lerin toplam sayilarini yazdir ve tum cell'leri tek tek yazdir
    @Test
    public void printCells(){
        System.out.println(" Print Cells");
       List<WebElement> tumCells = Driver.getDriver().findElements(By.xpath("//tbody//td"));
       //table body'dek' cells'lerin toplam sayilarini yazdir
        System.out.println("Toplam cell saysi => "+ tumCells.size());

        //tum cell'leri tek tek yazdir
        int cellNumarasi = 1;
        for (WebElement herCell : tumCells){
            System.out.println(cellNumarasi+ " : "+ herCell.getText());
            cellNumarasi++;
        }

        Driver.closeDriver();
    }
    //    test method: printColumns() ve columns'larin toplam sayisini yazdir ve 5. column yazdir
    @Test
    public void printColumns(){
        //columns'larin toplam sayisini yazdir
        List<WebElement> tumHeader = Driver.getDriver().findElements(By.xpath("//th"));
        int ColumnSayisi = tumHeader.size();
        System.out.println(ColumnSayisi+ "Column vardir");

       // 5. column yazdir
       List<WebElement> column5Data= Driver.getDriver().findElements(By.xpath("//table//tbody//tr//td[5]"));

       int hucreNumarasi =1;
       for (WebElement herHucre : column5Data){
           System.out.println("Hucre "+hucreNumarasi+ " : "+herHucre.getText());
           hucreNumarasi++;
       }

    }
    ////    test method: printData(int row, int column); Bu method veriline cell'i yazdirmali
    @Test
    public void printData(int row, int column){

        //ornek : printData(2, 3); 2. row ve 3 column yazdirmali

    }

}
