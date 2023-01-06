package com.excelautomation;

import com.utilities.ExcelUtil;
import org.apache.poi.ss.usermodel.*;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class Day12_C12_ReadExcel {
        /*
    Import the apache poi dependency in your pom file
    resources package olustur > java altinda acilmali (java'ya sag tikla ve dosyayi olustur)
    Add the excel file on the resources folder
    Yeni package olustur: excelautomation
    Yeni class olustur : ReadExcel
    test method olustur: readExcel()
    Dosyanin adresini String olarak bir konteynira koy
    dosyayi ac
    fileinputstream kullanarak workbook'u ac
    ilk worksheet'i ac
    ilk row'a git
    ilk row'daki ilk cell'e git ve yazdir
    ilk row'daki ikinci cell'e git ve yazdir
    2nd row'daki ilk cell'e git ve datanin ABD'ye esit oldugunu assert e
    3rd row'daki 2nd cell-chain the row and cell
    row sayisini bul
    Kullanilan row sayisini bul
    Ulke ve baskent key-value ciftlerini map object olarak yazdir
    */

    @Test
    public void readExcel() throws IOException {
//Dosyanin adresini String olarak bir konteynira koy
        String path = "./src/test/java/resources/Baskent.xlsx";
//        dosyayi ac
        FileInputStream fileInputStream = new FileInputStream(path);
 //fileinputstream kullanarak workbook'u ac;
       Workbook workbook = WorkbookFactory.create(fileInputStream);
   //ilk worksheet'i ac
       Sheet sheet =  workbook.getSheetAt(0); //sheet sayfalari '0.' index'ten baslar
   //ilk row'a git
        Row ilkRow =sheet.getRow(0);// row'lar index o'dan baslar

   //ilk row'daki ilk cell'e git ve yazdir
      Cell ilkCell =  ilkRow.getCell(0);// cell indexi 'o'dan baslar
        System.out.println(ilkCell);

        // 2. row'daki ikinci cell'e git ve yazdir
        //1. yol
//       Row row2 = sheet.getRow(1);
//       Cell cell21= row2.getCell(1);
//        System.out.println(cell21);//DC
        //2.yol
        Cell row1Cell2 = sheet.getRow(0).getCell(1);
        System.out.println(row1Cell2);//Baskent

//2nd row'daki ilk cell'e git ve datanin ABD'ye esit oldugunu assert e
        Cell row2Cell1 = sheet.getRow(1).getCell(0);
        System.out.println("Data ABD olmali : " +row2Cell1);

        boolean esitMi= row2Cell1.toString().equals("ABD");
        Assert.assertTrue(esitMi);

 //3rd row'daki 2nd cell
        Cell row3Cell2 = sheet.getRow(2).getCell(1);
        System.out.println(row3Cell2);//Paris
//row sayisini bul
        int rowSayisi= sheet.getLastRowNum()+ 1;
        System.out.println(rowSayisi);//11
// Kullanilan row sayisini bul
        int kullanilanRowSayisi = sheet.getPhysicalNumberOfRows();
        System.out.println(kullanilanRowSayisi);

//Ulke ve baskent key-value ciftlerini map object olarak yazdir

        Map<String, String> dunyaBasketleri = new HashMap<>();

        int ulkeColumn =0;
        int baskentColumn =1;

     /*
     row numarasi 1'den baslar, cunku header 0. indextedir
     En sondaki row'un indeksi : lastRowNumber yada sheet.getLastRowNum()+ 1;

     ABD        : sheet. getRow(1). getCEll(0);
     Fransa     : sheet. getRow(2). getCEll(0);
     .....
      */
      for(int rowNumarasi=1; rowNumarasi<rowSayisi; rowNumarasi++ ){
         String ulke = sheet.getRow(rowNumarasi).getCell(ulkeColumn).toString();
         String baskent =   sheet.getRow(rowNumarasi).getCell(baskentColumn).toString();

          System.out.println("Ulke: " + ulke);
          System.out.println("Baskent: " +baskent);

          dunyaBasketleri.put(ulke, baskent);// Map'e ulke ve baskent eklemis olduk

      }
        System.out.println(dunyaBasketleri);

    }

    @Test
    public void excelUtilDemo(){
        String path = "./src/test/java/resources/Baskent.xlsx";
        String sheetName ="Sayfa1";

        // //ExcelUtil class'i okumak icin once bir nesne olusturduk
        ExcelUtil excelUtil = new ExcelUtil(path, sheetName);
        //ExcelUtil'de methodlari cagirabiliriz

        System.out.println(excelUtil.getDataList());

        System.out.println(excelUtil.columnCount());//2

        System.out.println(excelUtil.rowCount());//11

        System.out.println(excelUtil.getCellData(5, 1));//Ottowa

        System.out.println(excelUtil.getColumnsNames());//[Ulke, Baskent]


    }

}
