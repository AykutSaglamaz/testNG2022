package com.tests;

import com.utilities.Driver;
import org.testng.Assert;
import org.testng.annotations.Test;


public class Day09_C4_FirstDriverTest {
    /*
    Amazon'a git
    title'in 'Amazon'u icerdigini verify et-
    Driver class calistigini check et
     */
    @Test
    public void amazonTitleTest(){
        //Amazon'a git
        Driver.getDriver().get("https://www.amazon.com");

//        title'in 'Amazon'u icerdigini verify et-
        String actualTitle = Driver.getDriver().getTitle();
        System.out.println(Driver.getDriver().getTitle());
//        Driver class calistigini check et
        Assert.assertTrue(actualTitle.contains("Amazon"));


        Driver.closeDriver();
    }

}
