package com.example.exceleditor;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Array;

public class Excel_Commands {

    public static HSSFSheet takeFile(String sheetName , InputStream FileLocation) throws IOException {
        InputStream inp = FileLocation;
        HSSFWorkbook wb = new HSSFWorkbook(new POIFSFileSystem(inp));
        HSSFSheet sheet = wb.getSheet(sheetName);
        return sheet ;
    }

    public static HSSFWorkbook takeFileReturnWorkbook(String sheetName , String FileLocation) throws IOException {
        InputStream inp = new FileInputStream(FileLocation);
        HSSFWorkbook wb = new HSSFWorkbook(new POIFSFileSystem(inp));
        HSSFSheet sheet = wb.getSheet(sheetName);
        return wb ;
    }

    public static String getReferenceNumber(HSSFSheet sheetForRef){
        CellReference referenceNumberCell = new CellReference("I6");
        Row referenceNumberRow = sheetForRef.getRow(referenceNumberCell.getRow());
        Cell referenceNumber = referenceNumberRow.getCell(referenceNumberCell.getCol());
        System.out.println("Inside Method" + referenceNumber.getStringCellValue());
        return referenceNumber.toString() ;
    }

    public static String getDate(HSSFSheet sheetForDate){
        CellReference dateCell = new CellReference("I5");
        Row dateRow = sheetForDate.getRow(dateCell.getRow());
        Cell date = dateRow.getCell(dateCell.getCol());
        System.out.println(date);
        return date.toString();
    }

    public static String[] getAccNumber(int accNumberNumberOfLots, HSSFSheet sheetForAccNumber){

        String[] accArr = new String[accNumberNumberOfLots] ;
        for(int i=0 ; i<accNumberNumberOfLots ;i++ ){
            int upp = 14 + i ;
            CellReference accCell = new CellReference("F"+upp);
            Row accRow = sheetForAccNumber.getRow(accCell.getRow());
            Cell account = accRow.getCell(accCell.getCol());
            accArr[i] = account.toString();
        }
        System.out.println(java.util.Arrays.toString(accArr));
        return accArr;
    }

    public static String[] getAccName(int accNameNumberOfLots , HSSFSheet sheetForName){
        String[] nameArr = new String[accNameNumberOfLots] ;
        for(int i=0 ; i<accNameNumberOfLots ;i++ ){
            int upp = 14 + i ;
            CellReference nameCell = new CellReference("G"+upp);
            Row nameRow = sheetForName.getRow(nameCell.getRow());
            Cell name = nameRow.getCell(nameCell.getCol());
            nameArr[i] = name.toString();
        }
        return nameArr;
    }

    public static String[] getAmmDeposite(int ammDepositeNumberOfLots,HSSFSheet sheetForAmmDeposite){
        String[] RDArr = new String[ammDepositeNumberOfLots] ;
        for(int i=0 ; i<ammDepositeNumberOfLots ;i++ ){
            int upp = 14 + i ;
            CellReference RDCell = new CellReference("H"+upp);
            Row RDRow = sheetForAmmDeposite.getRow(RDCell.getRow());
            Cell RD = RDRow.getCell(RDCell.getCol());
            RDArr[i] = RD.toString();
        }
        System.out.println(java.util.Arrays.toString(RDArr));
        return RDArr;
    }

    public static String[] getNumOfInstal(int numOfInstallNumberOfLots , HSSFSheet sheetForNumOFInstall){
        String[] numArr = new String[numOfInstallNumberOfLots] ;
        for(int i=0 ; i<numOfInstallNumberOfLots ;i++ ){
            int upp = 14 + i ;
            CellReference numCell = new CellReference("M"+upp);
            Row numRow = sheetForNumOFInstall.getRow(numCell.getRow());
            Cell num = numRow.getCell(numCell.getCol());
            numArr[i] = num.toString();
        }
        System.out.println(java.util.Arrays.toString(numArr));
        return numArr;
    }

    public static String[] getDefaults(int defaultsNumberOfLots , HSSFSheet sheetForDefaults){
        String[] defArr = new String[defaultsNumberOfLots] ;
        for(int i=0 ; i<defaultsNumberOfLots ;i++ ){
            int upp = 14 + i ;
            CellReference defCell = new CellReference("O"+upp);
            Row defRow = sheetForDefaults.getRow(defCell.getRow());
            Cell def = defRow.getCell(defCell.getCol());
            defArr[i] = def.toString();
        }
        System.out.println(java.util.Arrays.toString(defArr));
        return defArr;
    }

    public static String getTotalAmm(HSSFSheet sheetForTotalAmm , int numberofLots){
        int ammNo = numberofLots + 15 ;
        CellReference ammCell = new CellReference("J"+ammNo);
        Row ammRow = sheetForTotalAmm.getRow(ammCell.getRow());
        Cell amm = ammRow.getCell(ammCell.getCol());
        System.out.println(amm);
        return amm.toString() ;
    }

    public static void savingTheFile(String referenceNumberex , HSSFWorkbook workbooksv) throws IOException {
        FileOutputStream out = new FileOutputStream(referenceNumberex+".xls");
        workbooksv.write(out);
    }

    public static void insertRows(int insertRowsNumOfLots , HSSFSheet tempSheetForInsertRows){
        int numberOfLots = insertRowsNumOfLots-1 ;
        tempSheetForInsertRows.shiftRows(7,8,numberOfLots);
    }

    public static void setName(HSSFSheet sheet , int LotNum , String[] nameArr ,CellStyle style){

//                 Inserting account number data

        for(int i =0 ; i<LotNum ; i++){
            int tempAccNum = 7+i;
            CellReference nameCell = new CellReference("C"+tempAccNum);
            Row nameRow = sheet.getRow(nameCell.getRow());
            Cell name =  nameRow.createCell(nameCell.getCol());
            String nameOf = (String) Array.get(nameArr, i );
            System.out.println(nameOf);
            name.setCellValue(nameOf);
            name.setCellStyle(style);
        }

    }

    public static void setRefNum(HSSFSheet sheetSetRef , String referenceValue , int NumberOfLots){

        for(int i =0 ; i<NumberOfLots ; i++){

            int tempAccNum = 7+i;
            CellReference tempAccRefCell = new CellReference("A"+tempAccNum);
            Row tempAccRefRow = sheetSetRef.createRow(tempAccRefCell.getRow());
            Cell tempAccRef =  tempAccRefRow.createCell(tempAccRefCell.getCol());
            tempAccRefRow.setHeightInPoints(24);
            tempAccRef.setCellValue(referenceValue);
//            row.getCell(sass.getCol()).setCellStyle(style);

        }

    }

    public static void setRefNumMain(HSSFSheet sheetForRef , String Ref){
        CellReference tempDateCell =  new CellReference("D5");
        Row tempDateRow = sheetForRef.getRow(tempDateCell.getRow());
        Cell tempDate = tempDateRow.getCell(tempDateCell.getCol());
        tempDate.setCellValue(Ref);
    }

    public static void setDate(HSSFSheet sheetForDate , String date){
        CellReference tempDateCell =  new CellReference("D4");
        Row tempDateRow = sheetForDate.getRow(tempDateCell.getRow());
        Cell tempDate = tempDateRow.getCell(tempDateCell.getCol());
        tempDate.setCellValue(date);
    }

    public static void setAccNum(HSSFSheet sheet , int LotNum , String[] nameArr){

//                 Inserting account number data

        for(int i =0 ; i<LotNum ; i++){
            int tempAccNum = 7+i;
            CellReference nameCell = new CellReference("B"+tempAccNum);
            Row nameRow = sheet.getRow(nameCell.getRow());
            Cell name =  nameRow.createCell(nameCell.getCol());
            String nameOf = (String) Array.get(nameArr, i );
            System.out.println(nameOf);
            name.setCellValue(nameOf);

        }

    }

    public static void setRdDemon(HSSFSheet sheet , int LotNum , String[] nameArr){

//                 Inserting account number data

        for(int i =0 ; i<LotNum ; i++){
            int tempAccNum = 7+i;
            CellReference nameCell = new CellReference("D"+tempAccNum);
            Row nameRow = sheet.getRow(nameCell.getRow());
            Cell name =  nameRow.createCell(nameCell.getCol());
            String nameOf = (String) Array.get(nameArr, i );
            System.out.println(nameOf);
            name.setCellValue(nameOf);

        }

    }

    public static void setDepsiteAmm(HSSFSheet sheet , int LotNum , String[] nameArr){

//                 Inserting account number data

        for(int i =0 ; i<LotNum ; i++){
            int tempAccNum = 7+i;
            CellReference nameCell = new CellReference("E"+tempAccNum);
            Row nameRow = sheet.getRow(nameCell.getRow());
            Cell name =  nameRow.createCell(nameCell.getCol());
            String nameOf = (String) Array.get(nameArr, i );
            System.out.println(nameOf);
            name.setCellValue(nameOf);

        }

    }

    public static void setNumOfInstall(HSSFSheet sheet , int LotNum , String[] nameArr){

//                 Inserting account number data

        for(int i =0 ; i<LotNum ; i++){
            int tempAccNum = 7+i;
            CellReference nameCell = new CellReference("F"+tempAccNum);
            Row nameRow = sheet.getRow(nameCell.getRow());
            Cell name =  nameRow.createCell(nameCell.getCol());
            String nameOf = (String) Array.get(nameArr, i );
            System.out.println(nameOf);
            name.setCellValue(nameOf);

        }

    }

    public static void setPenalty(HSSFSheet sheet , int LotNum , String[] nameArr){

//                 Inserting account number data

        for(int i =0 ; i<LotNum ; i++){
            int tempAccNum = 7+i;
            CellReference nameCell = new CellReference("G"+tempAccNum);
            Row nameRow = sheet.getRow(nameCell.getRow());
            Cell name =  nameRow.createCell(nameCell.getCol());
            String nameOf = (String) Array.get(nameArr, i );
            System.out.println(nameOf);
            name.setCellValue(nameOf);

        }

    }

    public static void setRefNumEnd(HSSFSheet sheetForRef , String Ref , int numberOfLots , CellStyle style){
        int amm = numberOfLots+8;
        CellReference tempDateCell =  new CellReference("A"+amm);
        Row tempDateRow = sheetForRef.getRow(tempDateCell.getRow());
        Cell tempDate = tempDateRow.getCell(tempDateCell.getCol());
        tempDate.setCellValue(Ref);

        //Styles

        tempDateRow.getCell(tempDateCell.getCol()).setCellStyle(style);
    }

    public static void setAmmEnd(HSSFSheet sheetForRef , String Ref , int numberOfLots ,CellStyle style){
        int amm = numberOfLots+8;
        CellReference tempDateCell =  new CellReference("E"+amm);
        Row tempDateRow = sheetForRef.getRow(tempDateCell.getRow());
        Cell tempDate = tempDateRow.getCell(tempDateCell.getCol());
        tempDate.setCellValue(Ref);

        //Styles

        tempDateRow.getCell(tempDateCell.getCol()).setCellStyle(style);
    }
}
