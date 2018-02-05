package pdftoexcelmacrowriter;

import java.awt.Choice;
import javax.swing.*;
import java.awt.event.*;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JFileChooser;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

class RadioButtonExample extends JFrame implements ActionListener {

    int returnValue = 0;
    JRadioButton rb1, rb2;
    JButton b, close;
    ButtonGroup bg;
    Choice district;
    JLabel lable;
    String[] Districts = {"1-BELGAUM", "2-BAGALKOT", "3-BIJAPUR", "4-GULBARGA", "5-BIDAR", "6-RAICHUR", "7- KOPPAL", "8-GADAG",
 "9-DHARWAD", "10-UTTARA KANNADA", "11-HAVERI", "12-BELLARY", "13-CHITRADURGA", "14-DAVANGERE", "15- SHIMOGA",
 "16-UDUPI", "17-CHIKMAGALUR", "18-TUMKUR", "19-CHIKKABALLAPUR", "20-KOLAR", "21-BANGALORE",
 "22-BANGALORE RURAL", "23-RAMANAGARAM", "24-MANDYA", "25-HASSAN", "26-DAKSHINA KANNADA",
 "27-KODAGU", "28-MYSORE", "29-CHAMARAJNAGAR", "35-YADGIR"};
    RadioButtonExample() {
        rb1 = new JRadioButton("English PDF File");
        rb1.setBounds(50, 50, 150, 30);
        rb2 = new JRadioButton("Kannada PDF File");
        rb2.setBounds(50, 100, 150, 30);
        bg = new ButtonGroup();
        bg.add(rb1);
        bg.add(rb2);
        lable = new JLabel("Select District");
        lable.setBounds(50, 150, 80, 30);
        district = new Choice();
        
        district.setBounds(150, 150, 110, 30);
        for (String District : Districts) {
            district.add(District);
        }

        b = new JButton("Ok");
        b.setBounds(50, 200, 80, 30);
        b.addActionListener(this);
        close = new JButton("Close");
        close.setBounds(150, 200, 80, 30);
        close.addActionListener(this);

        add(rb1);
        add(rb2);
        add(b);
        add(close);
        add(district);
        add(lable);
        setSize(300, 300);
        
        setLayout(null);
        setVisible(true);
    }

    public void actionPerformed(ActionEvent e) {
        if (e.getSource() == b) {
            if (rb2.isSelected() || rb1.isSelected()) {
                JFileChooser chooser = new JFileChooser();
                chooser.setCurrentDirectory(new java.io.File("."));
                chooser.setDialogTitle("Select pdf File");
                chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
                chooser.setAcceptAllFileFilterUsed(true);

                Runtime rt = Runtime.getRuntime();
                if (chooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
                    String command = "java -jar lib\\pdfbox-app-2.0.8.jar ExtractText \"" + chooser.getSelectedFile() + "\" voterlist.txt";
                    System.out.println(command);
                    Process proc;
                    try {
                        proc = rt.exec(command);
                        System.out.println("Done PDF to txt");

                        BufferedReader stdError = new BufferedReader(new InputStreamReader(proc.getErrorStream()));

                        // read any errors from the attempted command
                        System.out.println("Here is the standard error of the command (if any):\n");
                        String s;
                        while ((s = stdError.readLine()) != null) {
                            System.out.println(s);
                        }
                    } catch (IOException ex) {
                        Logger.getLogger(RadioButtonExample.class.getName()).log(Level.SEVERE, null, ex);
                        
                    }

                    String TXTfilename = "voterlist.txt";
                    BufferedReader reader = null;
                    try {
                        reader = new BufferedReader(new FileReader(TXTfilename));
                    } catch (FileNotFoundException ex) {
                        Logger.getLogger(RadioButtonExample.class.getName()).log(Level.SEVERE, null, ex);
                    }

                    XSSFWorkbook workbook = new XSSFWorkbook();
                    XSSFSheet sheet = workbook.createSheet("Voter List");

                    String iMacroInCsvFileName = chooser.getSelectedFile().getName();
                    iMacroInCsvFileName = iMacroInCsvFileName.replace(".pdf", "_Input2Macro.csv");
                    //iMacroInCsvFileName = iMacroInCsvFileName.replaceAll(".PDF", "_Input2Macro.csv");
                    BufferedWriter csvWriter = null;
                    try {
                        csvWriter = new BufferedWriter(new FileWriter(iMacroInCsvFileName));
                    } catch (IOException ex) {
                        Logger.getLogger(RadioButtonExample.class.getName()).log(Level.SEVERE, null, ex);
                    }

                    String iMacroOutCsvFileName = chooser.getSelectedFile().getName();
                    iMacroOutCsvFileName = iMacroOutCsvFileName.replace(".pdf", "_OutputFromMacro.csv");
                    //iMacroOutCsvFileName = iMacroOutCsvFileName.replaceAll(".PDF", "_OutputFromMacro.csv");

                    String MacroFileName = chooser.getSelectedFile().getName();
                    MacroFileName = MacroFileName.replace(".pdf", "_MacroFile.iim");
                    //MacroFileName = MacroFileName.replaceAll(".PDF", "_MacroFile.iim");
                    try {
                        BufferedWriter writer = new BufferedWriter(new FileWriter(MacroFileName));
                        writer.write("VERSION BUILD=9030808 RECORDER=FX\n");

                        writer.write("TAB T=1\n");
                        writer.write("TAB CLOSEALLOTHERS\n\n");

                        writer.write("SET !TIMEOUT_PAGE 5\n");
                        writer.write("SET !ERRORIGNORE YES\n");
                        writer.write("SET !EXTRACT_TEST_POPUP NO\n\n");

                        writer.write("SET !DATASOURCE " + iMacroInCsvFileName + "\n");
                        writer.write("SET !LOOP 1\n");

                        writer.write("SET !DATASOURCE_LINE {{!LOOP}}\n");
                        writer.write("SET !TIMEOUT_PAGE 7\n");
                        writer.write("URL GOTO=http://ceokarnataka.kar.nic.in/SearchWithEpicNo_New.aspx#\n\n");
                        writer.write("TAG POS=1 TYPE=SELECT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_ddlDistrict CONTENT=%"+(district.getSelectedIndex()+1)+"\n");
                        writer.write("TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtEpic CONTENT={{!COL2}}\n");
                        writer.write("SET !TIMEOUT_PAGE 2\n");
                        writer.write("TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_btnSearch\n\n");
                    

                        writer.write("TAG POS=1 TYPE=INPUT:TEXT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_txtimgcode CONTENT=abcde\n");
                        writer.write("TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_Button1\n");
                        writer.write("TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:aspnetForm ATTR=ID:ctl00_ContentPlaceHolder1_GridView1_ctl02_btnDetails\n\n");

                        writer.write("TAG POS=2 TYPE=TABLE ATTR=TXT:* EXTRACT=TXT\n");
                        writer.write("SAVEAS TYPE=EXTRACT FOLDER=* FILE=" + iMacroOutCsvFileName + "\n\n");

                        writer.write("TAG POS=3 TYPE=TABLE ATTR=TXT:* EXTRACT=TXT\n");
                        writer.write("SAVEAS TYPE=EXTRACT FOLDER=* FILE=" + iMacroOutCsvFileName + "\n\n");

                        writer.close();
                    } catch (IOException ex) {
                        Logger.getLogger(RadioButtonExample.class.getName()).log(Level.SEVERE, null, ex);
                    }

                    String line;

                    voterDetails voter = new voterDetails();
                    int rowCount = 0;
                    int columnCount = 0;
                    Row row;
                    Cell cell;

                    if (rb1.isSelected()) {
                        String line1[] = new String[8];
                        String previousLine = "";

                        row = sheet.createRow(rowCount++);
                        cell = row.createCell(columnCount++);
                        cell.setCellValue("SlNo.");
                        cell = row.createCell(columnCount++);
                        cell.setCellValue("IDCardNo");
                        cell = row.createCell(columnCount++);
                        cell.setCellValue("Name");
                        cell = row.createCell(columnCount++);
                        cell.setCellValue("Fathers Name");
                        cell = row.createCell(columnCount++);
                        cell.setCellValue("Sex");
                        cell = row.createCell(columnCount++);
                        cell.setCellValue("Age");
                        cell = row.createCell(columnCount++);
                        cell.setCellValue("HouseNo");
                        try {
                            while ((line = reader.readLine()) != null) {

                                System.out.println(line);
                                if (line.isEmpty()) {
                                    continue;
                                }

                                if (line.startsWith("Name :")) {

                                    int i = 0;
                                    while (i < 6) {

                                        if ((line1[i++] = reader.readLine()) == null) {
                                            break;
                                        }

                                    }

                                    if(line1[2].contains("Photo"))
                                    {
                                        line1[2] = line1[5];
                                        line1[3] = reader.readLine();
                                        line1[4] = reader.readLine();
                                        line1[5] = reader.readLine();
                                        
                                    }
                                    
                                    line1[0] = line1[0].replace("House No.:", "");
                                    voter.setHouseNo(line1[0]);
                                    voter.setRName(line1[1]);
                                    voter.setEName(line1[2]);
                                    voter.setIDCardNo(line1[3].trim());

                                    if (line1[4].startsWith("Sex: FemaleAge: ")) {

                                        voter.setSex("Female");
                                        line1[4] = line1[4].replace("Sex: FemaleAge: ", "");
                                        voter.setAge(line1[4]);

                                    } else if (line1[4].startsWith("Sex: MaleAge: ")) {
                                        voter.setSex("Male");
                                        line1[4] = line1[4].replace("Sex: MaleAge: ", "");
                                        voter.setAge(line1[4]);
                                    }

                                    row = sheet.createRow(rowCount++);
                                    columnCount = 0;
                                    cell = row.createCell(columnCount++);
                                    cell.setCellValue(previousLine);
                                    cell = row.createCell(columnCount++);
                                    cell.setCellValue(voter.getIDCardNo());
                                    cell = row.createCell(columnCount++);
                                    cell.setCellValue(voter.getEName());
                                    cell = row.createCell(columnCount++);
                                    cell.setCellValue(voter.getRName());
                                    cell = row.createCell(columnCount++);
                                    cell.setCellValue(voter.getSex());
                                    cell = row.createCell(columnCount++);
                                    cell.setCellValue(voter.getAge());
                                    cell = row.createCell(columnCount++);
                                    cell.setCellValue(voter.getHouseNo());

                                    csvWriter.write(previousLine.trim() + ",");

                                    csvWriter.write(voter.getIDCardNo().trim() + "\n");

                                }

                                previousLine = line;

                            }
                        } catch (IOException ex) {
                            Logger.getLogger(RadioButtonExample.class.getName()).log(Level.SEVERE, null, ex);
                        }

                    }
                    if (rb2.isSelected()) {

                        row = sheet.createRow(rowCount++);
                        cell = row.createCell(columnCount++);
                        cell.setCellValue("SlNo.");
                        cell = row.createCell(columnCount++);
                        cell.setCellValue("IDCardNo");
                            int ii =0;
                        try {
                            while ((line = reader.readLine()) != null) {
                                System.out.println(line);
                                line = line.replace("# ", "");
                                if (line.isEmpty()) {
                                    continue;
                                }

                                String[] splitStr = line.trim().split("\\s+");
                                if (splitStr.length < 6 || ii++<1) {
                                    continue;
                                }
                                row = sheet.createRow(rowCount++);
                                cell = row.createCell(0);
                                cell.setCellValue(splitStr[0]);
                                cell = row.createCell(1);
                                cell.setCellValue(splitStr[1]);

                                row = sheet.createRow(rowCount++);
                                cell = row.createCell(0);
                                cell.setCellValue(splitStr[2]);
                                cell = row.createCell(1);
                                cell.setCellValue(splitStr[3]);

                                row = sheet.createRow(rowCount++);
                                cell = row.createCell(0);
                                cell.setCellValue(splitStr[4]);
                                cell = row.createCell(1);
                                cell.setCellValue(splitStr[5]);

                                csvWriter.write(splitStr[0] + "," + splitStr[1] + "\n");
                                csvWriter.write(splitStr[2] + "," + splitStr[3] + "\n");
                                csvWriter.write(splitStr[4] + "," + splitStr[5] + "\n");

                            }
                        } catch (IOException ex) {
                            Logger.getLogger(RadioButtonExample.class.getName()).log(Level.SEVERE, null, ex);
                        }

                    }

                    try {
                        reader.close();
                        csvWriter.close();
                    } catch (IOException ex) {
                        Logger.getLogger(RadioButtonExample.class.getName()).log(Level.SEVERE, null, ex);
                    }

                    String excelFile = chooser.getSelectedFile().getName();
                    excelFile = excelFile.replace(".pdf", ".xlsx");
                    //excelFile = excelFile.replaceAll(".PDF", ".xlsx");

                    try {
                        FileOutputStream outputStream = new FileOutputStream(excelFile);
                        workbook.write(outputStream);
                    } catch (FileNotFoundException ex) {
                        Logger.getLogger(RadioButtonExample.class.getName()).log(Level.SEVERE, null, ex);
                    } catch (IOException ex) {
                        Logger.getLogger(RadioButtonExample.class.getName()).log(Level.SEVERE, null, ex);
                    }
                    System.out.println("Excel File: " + excelFile);
                    System.out.println("iMacro File: " + MacroFileName);
                    System.out.println("Input for iMacro File: " + iMacroInCsvFileName);

                    JOptionPane.showMessageDialog(this, "Excel File: " + excelFile + "\niMacro File: " + MacroFileName + "\nInput for iMacro File: " + iMacroInCsvFileName);

                } else {
                    JOptionPane.showMessageDialog(this, "No file selected");
                }
            } else {
                JOptionPane.showMessageDialog(this, "No file Type Selected");
            }
        }

        if (e.getSource() == close) {
            this.dispose();
        }
    }
}
