/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package csv2excel;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import javax.lang.model.element.Element;
import javax.lang.model.util.Elements;
import javax.swing.JFileChooser;
import javax.swing.text.Document;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import org.jsoup.Jsoup;

/**
 * A very simple program that writes some data to an Excel file using the Apache
 * POI library.
 *
 * @author www.codejava.net
 *
 */
public class Csv2Excel {

    public static void main(String[] args) throws IOException {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Voterlist");
        int rowCount = 0;
        int columnCount = 0;
        Row row1;
        Cell cell;

        row1 = sheet.createRow(rowCount);
        columnCount = 0;
        cell = row1.createCell(columnCount++);
        cell.setCellValue("ACNO");
        cell = row1.createCell(columnCount++);
        cell.setCellValue("EAssemblyConstituency");
        cell = row1.createCell(columnCount++);
        cell.setCellValue("KAssemblyConstituency");
        cell = row1.createCell(columnCount++);
        cell.setCellValue("PartNo");
        cell = row1.createCell(columnCount++);
        cell.setCellValue("PollingStation");
        cell = row1.createCell(columnCount++);
        cell.setCellValue("SlnNo");
        cell = row1.createCell(columnCount++);
        cell.setCellValue("ESection");
        cell = row1.createCell(columnCount++);
        cell.setCellValue("KSection");
        cell = row1.createCell(columnCount++);
        cell.setCellValue("EFirstName");
        cell = row1.createCell(columnCount++);
        cell.setCellValue("KFirstName");
        cell = row1.createCell(columnCount++);
        cell.setCellValue("ELastName");
        cell = row1.createCell(columnCount++);
        cell.setCellValue("KLastName");
        cell = row1.createCell(columnCount++);
        cell.setCellValue("ERelationFirstName");
        cell = row1.createCell(columnCount++);
        cell.setCellValue("KRelationFirstName");
        cell = row1.createCell(columnCount++);
        cell.setCellValue("ERelationLastName");
        cell = row1.createCell(columnCount++);
        cell.setCellValue("KRelationLastName");
        cell = row1.createCell(columnCount++);
        cell.setCellValue("sex");
        cell = row1.createCell(columnCount++);
        cell.setCellValue("age");
        cell = row1.createCell(columnCount++);
        cell.setCellValue("HouseNo");
        cell = row1.createCell(columnCount++);
        cell.setCellValue("IDCardNo");
        cell = row1.createCell(columnCount++);
        cell.setCellValue("OldIDCardNo");

        
        JFileChooser chooser = new JFileChooser();
        chooser.setCurrentDirectory(new java.io.File("."));
        chooser.setDialogTitle("Select pdf File");
        chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
        chooser.setAcceptAllFileFilterUsed(true);

        //       Runtime rt = Runtime.getRuntime();
        if (chooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
            System.out.println("getCurrentDirectory(): " + chooser.getCurrentDirectory());
            System.out.println("getSelectedFile() : " + chooser.getSelectedFile());
            File folder = chooser.getSelectedFile();
            System.out.println("Folder Selected = : " + folder.getAbsolutePath());

            voterDetails voter = new voterDetails();

            BufferedReader reader = new BufferedReader(new InputStreamReader( new FileInputStream(folder.getAbsolutePath()), "UTF8"));
            //BufferedReader reader = new BufferedReader(new FileReader(folder.getAbsolutePath()));
            String line = "";
            while ((line = reader.readLine()) != null) {

                if (line.contains("\"\",\"ACNO\",\"AssemblyConstituency\",")) {

                                        
                    System.out.println(line);
                    String[] line1 = new String[16];

                    int loop = 0;
                    while (loop < 16) {
                        if((line1[loop] = reader.readLine()) == null)
                        {
                            break;
                        }
                        System.out.println(line1[loop]);
                        loop++;
                    }
                    String parameter;
                    String[] parameter1 = new String[5];
                    
                    System.out.println("ACNO: "+line1[0]);
                    voter.setACNO(line1[0].substring(4, 6).trim());
                    System.out.println("ACNO: "+voter.getACNO());
                    
                    System.out.println("AssemblyConstituency: "+line1[2]);
                    parameter = line1[2];
                    parameter = parameter.substring(45);
                    parameter = parameter.replace("\"", "");
                    parameter1 = parameter.split("/");
                    voter.setKAssemblyConstituency(parameter1[0].trim());
                    voter.setEAssemblyConstituency(parameter1[1].trim());
                    System.out.println("KAssemblyConstituency: "+voter.getKAssemblyConstituency());
                    System.out.println("EAssemblyConstituency: "+voter.getEAssemblyConstituency());
                    
                    System.out.println("PartNo: "+line1[3]);
                    parameter = line1[3];
                    parameter = parameter.substring(29);
                    parameter = parameter.replace("\"", "");
                    voter.setPartNo(parameter.trim());
                    System.out.println("PartNo: "+voter.getPartNo());

                    System.out.println("PollingStation: "+line1[4]);
                    parameter = line1[4];
                    parameter = parameter.substring(29);
                    parameter = parameter.replace("\"", "");
                    voter.setPollingStation(parameter.trim());
                    System.out.println("PollingStation: "+voter.getPollingStation());

                    System.out.println("SlnNo: "+line1[5]);
                    parameter = line1[5];
                    parameter = parameter.substring(23);
                    parameter = parameter.replace("\"", "");
                    voter.setSlnNo(parameter.trim());
                    System.out.println("SlnNo: "+voter.getSlnNo());

                    System.out.println("Section: "+line1[6]);
                    parameter = line1[6];
                    parameter = parameter.substring(18);
                    parameter = parameter.replace("\"", "");
                    parameter1 = parameter.split("/");
                    voter.setKSection(parameter1[0].trim());
                    voter.setESection(parameter1[1].trim());
                    System.out.println("KSection: "+voter.getKSection());
                    System.out.println("ESection: "+voter.getESection());

                    System.out.println("FirstName: "+line1[7]);
                    parameter = line1[7];
                    parameter = parameter.substring(23);
                    parameter = parameter.replace("\"", "");
                    parameter1 = parameter.split("/");
                    voter.setKFirstName(parameter1[0].trim());
                    voter.setEFirstName(parameter1[1].trim());
                    System.out.println("KFirstName: "+voter.getKFirstName());
                    System.out.println("EFirstName: "+voter.getEFirstName());

                    System.out.println("LastName: "+line1[8]);
                    parameter = line1[8];
                    parameter = parameter.substring(22);
                    parameter = parameter.replace("\"", "");
                    parameter1 = parameter.split("/");
                    voter.setKLastName(parameter1[0].trim());
                    voter.setELastName(parameter1[1].trim());
                    System.out.println("KLastName: "+voter.getKLastName());
                    System.out.println("ELastName: "+voter.getELastName());

                    System.out.println("RelationFirstName: "+line1[9]);
                    parameter = line1[9];
                    parameter = parameter.substring(38);
                    parameter = parameter.replace("\"", "");
                    parameter1 = parameter.split("/");
                    voter.setKRelationFirstName(parameter1[0].trim());
                    voter.setERelationFirstName(parameter1[1].trim());
                    System.out.println("KRelationFirstName: "+voter.getKRelationFirstName());
                    System.out.println("ERelationFirstName: "+voter.getERelationFirstName());
                    /*
parameter = line1[10];
parameter = parameter.substring(38);
parameter = parameter.replace("\"", "");
parameter1 = parameter.split("/");
voter.setKRelationLastName(parameter1[0]);
voter.setERelationLastName(parameter1[1]);
                     */

                    System.out.println("Sex: "+line1[11]);
                    parameter = line1[11];
                    parameter = parameter.substring(14);
                    parameter = parameter.replace("\"", "");
                    voter.setSex(parameter.trim());
                    System.out.println("Sex: "+voter.getSex());

                    System.out.println("Age: "+line1[12]);
                    parameter = line1[12];
                    parameter = parameter.substring(16);
                    parameter = parameter.replace("\"", "");
                    voter.setAge(parameter.trim());
                    System.out.println("Age: "+voter.getAge());

                    System.out.println("HouseNo: "+line1[13]);
                    parameter = line1[13];
                    parameter = parameter.substring(25);
                    parameter = parameter.replace("\"", "");
                    voter.setHouseNo(parameter.trim());
                    System.out.println("HouseNo: "+voter.getHouseNo());

                    System.out.println("IDCardNo: "+line1[14]);
                    parameter = line1[14];
                    parameter = parameter.substring(34);
                    parameter = parameter.replace("\"", "");
                    voter.setIDCardNo(parameter.trim());
                    System.out.println("IDCardNo: "+voter.getIDCardNo());

                    
                    System.out.println("OldIDCardNo: "+line1[15]);
                    parameter = line1[15];
                    parameter = parameter.substring(43);
                    parameter = parameter.replace("\"", "");
                    voter.setOldIDCardNo(parameter.trim());
                    System.out.println("IDCardNo: "+voter.getOldIDCardNo());


                    /*   System.out.println("ACNO :"+voter.getACNO());
    System.out.println("PartNo :"+voter.getPartNo());
    System.out.println("SlnNo :"+voter.getSlnNo());
    System.out.println("EName :"+voter.getEName());
    System.out.println("VEName :"+voter.getVEName());
    System.out.println("VRName :"+voter.getVRName());
    System.out.println("RType :"+voter.getRType());
    System.out.println("Age :"+voter.getAge());
    System.out.println("IDCardNo :"+voter.getIDCardNo());
    System.out.println("HouseNo :"+voter.getHouseNo());
    System.out.println("SectionNo :"+voter.getSectionNo());
    System.out.println("VAddress :"+voter.getVAddress());
    System.out.println("VLocation :"+voter.getVLocation());
                     */

                    rowCount++;
                    row1 = sheet.createRow(rowCount);
                    columnCount = 0;
                    cell = row1.createCell(columnCount++);
                    cell.setCellValue(voter.getACNO());
                    cell = row1.createCell(columnCount++);
                    cell.setCellValue(voter.getEAssemblyConstituency());
                    cell = row1.createCell(columnCount++);
                    cell.setCellValue(voter.getKAssemblyConstituency());
                    cell = row1.createCell(columnCount++);
                    cell.setCellValue(voter.getPartNo());
                    cell = row1.createCell(columnCount++);
                    cell.setCellValue(voter.getPollingStation());
                    cell = row1.createCell(columnCount++);
                    cell.setCellValue(voter.getSlnNo());
                    cell = row1.createCell(columnCount++);
                    cell.setCellValue(voter.getESection());
                    cell = row1.createCell(columnCount++);
                    cell.setCellValue(voter.getKSection());
                    cell = row1.createCell(columnCount++);
                    cell.setCellValue(voter.getEFirstName());
                    cell = row1.createCell(columnCount++);
                    cell.setCellValue(voter.getKFirstName());
                    cell = row1.createCell(columnCount++);
                    cell.setCellValue(voter.getELastName());
                    cell = row1.createCell(columnCount++);
                    cell.setCellValue(voter.getKLastName());
                    cell = row1.createCell(columnCount++);
                    cell.setCellValue(voter.getERelationFirstName());
                    cell = row1.createCell(columnCount++);
                    cell.setCellValue(voter.getKRelationFirstName());
                    cell = row1.createCell(columnCount++);
                    cell.setCellValue(voter.getERelationLastName());
                    cell = row1.createCell(columnCount++);
                    cell.setCellValue(voter.getKRelationLastName());
                    cell = row1.createCell(columnCount++);
                    cell.setCellValue(voter.getSex());
                    cell = row1.createCell(columnCount++);
                    cell.setCellValue(voter.getAge());
                    cell = row1.createCell(columnCount++);
                    cell.setCellValue(voter.getHouseNo());
                    cell = row1.createCell(columnCount++);
                    cell.setCellValue(voter.getIDCardNo());
                    cell = row1.createCell(columnCount++);
                    cell.setCellValue(voter.getOldIDCardNo());
                }
            }

                    String xlsxFileName = chooser.getSelectedFile().getName();
        xlsxFileName = xlsxFileName.replace(".csv", "_Final.xlsx");
            try (FileOutputStream outputStream = new FileOutputStream(xlsxFileName))
            {
                workbook.write(outputStream);
                System.out.println("Final output file is generated : "+xlsxFileName);
            }
        }
    }
}
