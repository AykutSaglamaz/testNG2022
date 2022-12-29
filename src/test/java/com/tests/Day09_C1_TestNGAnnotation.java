package com.tests;

import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Ignore;
import org.testng.annotations.Test;

public class Day09_C1_TestNGAnnotation {
    /*
    @Test test caseleri olusturmak icin kullanilir TestNG
    @BeforeMEthod her bir @Test anotationdan once kosmasi icin kullanilir, Junit @Before ile aynidir
    @AfterMethod her bir @Test anotationdan sonra kosmasi icin kullanilir
    @Ignore bir @Test anotation- atlamak icin kullaniririr
    @Test(enabled = false)  durumunda @Test anotation- atlamak icin kullaniririr -- @Ignore gividir

    Test caseleri siralayabiliriz
    TestNG test caseleri alfabetik siraya gore kosturur- yukarin asagiya dogru calistirmaz

    Test case icin priority kullarak siralama yapabiliriz

   orn: @Test (priority =1)
     */

    @BeforeMethod
    public void setUp(){
        System.out.println("Before method kullandik");
    }

    @AfterMethod
    public void tearDown(){
        System.out.println("Afer method kullandik");
    }

    @Test (priority = 2)
    public void test6(){
        System.out.println("Test 6");
    }


    @Test (enabled = false)
    public void test2(){
        System.out.println("Test 2");
    }

    @Ignore
    @Test
    public void test3(){
        System.out.println("Test 3");
    }

    @Test (priority = -3)
    public void test4(){
        System.out.println("Test 4");
    }

    @Test (priority = -1)
    public void test5(){
        System.out.println("Test 5");
    }

    @Test (priority = 1)
    public void test7(){
        System.out.println("Test 7");
    }

}
