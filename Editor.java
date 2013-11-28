/*******************************************************************

Title: Java Program Created For Edward Mullings
Purpose: GUI style excel spreadsheet regex use.

Author: Xamboni

*******************************************************************/



package Editor;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import javax.swing.Box;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JRadioButton;
import javax.swing.JSeparator;
import javax.swing.JTextField;
import javax.swing.SwingConstants;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;



public class Editor implements ActionListener
{
	
	JTextField inputFilename = new JTextField(50);
	public static void main(String args[]) throws InvalidFormatException, IOException
	{
		
		Editor myEditor = new Editor();
    
	}
	
	public Editor() throws InvalidFormatException, IOException
	{
		int query = JOptionPane.showConfirmDialog(null, "You are about to edit an Excel spreadsheet.\n" +
				"You have the option to save a copy of the spreadsheet before edits are made.\nAre you sure you want to proceed?", null, JOptionPane.YES_NO_OPTION);
	
		boolean select = true;
	    JTextField column = new JTextField(5);
	    JTextField delimiter = new JTextField(5);	
	    JTextField rowCount = new JTextField(5);
	    JCheckBox copy = new JCheckBox("Create backup of original file?", select);
		
		if(query == JOptionPane.YES_OPTION)
			{
				
				//String inputFileName = "Carson Piping Cross Reference 10-9-2013 (Master).xls";
				try
				{
				    
				   // JPanel myPanel = new JPanel();
					
					
					Box myPanel = Box.createVerticalBox();
					
				    JButton fileChooser = new JButton("Choose File");
					myPanel.add(fileChooser); 
					myPanel.add(new JLabel("Input Filename:"));
				    myPanel.add(inputFilename);
				    fileChooser.addActionListener(this); 
				    //myPanel.add(Box.createVerticalStrut(200)); // a spacer
				    //myPanel.add(Box.createHorizontalStrut(15)); // a spacer
				    myPanel.add(new JLabel("Column: (Number of column to edit)"));
				    myPanel.add(column);
				    myPanel.add(Box.createHorizontalStrut(40)); // a spacer
				    myPanel.add(new JLabel("Delimiter: (Everything after in the cell will be deleted)"));
				    myPanel.add(delimiter);
				    myPanel.add(new JLabel("Row Count: (How many rows to change?"));
				    myPanel.add(rowCount);
				    myPanel.add(copy);
				    
				    int result = JOptionPane.showConfirmDialog(null, myPanel, 
				               "Please fill in the boxes.", JOptionPane.OK_CANCEL_OPTION);			    
				    
				    if (result == JOptionPane.OK_OPTION) 
				    {
				        if(inputFilename.getText().length() == 0)
				        {
					    	result = JOptionPane.showConfirmDialog(null, myPanel, 
						               "Filename was left blank. Exiting.", JOptionPane.OK_CANCEL_OPTION);
					    	System.exit(1);
				        }
				        
				        if(column.getText().length() == 0)
				        {
					    	result = JOptionPane.showConfirmDialog(null, myPanel, 
						               "Column was left blank. Exiting.", JOptionPane.OK_CANCEL_OPTION);
					    	System.exit(1);
				        }
				        
				        else if(delimiter.getText().length() == 0)
				        {
					    	result = JOptionPane.showConfirmDialog(null, myPanel, 
						               "Delimiter was left blank. Exiting.", JOptionPane.OK_CANCEL_OPTION);
					    	System.exit(1);
				        }
				    	
				        if(copy.isSelected()) 
				        {
							InputStream inp = new FileInputStream(inputFilename.getText());
						    Workbook wb = WorkbookFactory.create(inp);
					    	FileOutputStream fileOut = new FileOutputStream("Backup Copy " + inputFilename.getText());
					    	wb.write(fileOut);
						    fileOut.close();
				        }
				        	
						InputStream inp = new FileInputStream(inputFilename.getText());
					    Workbook wb = WorkbookFactory.create(inp);
					    org.apache.poi.ss.usermodel.Sheet sheet = wb.getSheetAt(0);
				        
					    for(int i = 0; i < Integer.parseInt(rowCount.getText()); i += 1)
					    {
					    	Row row = ((org.apache.poi.ss.usermodel.Sheet) sheet).getRow(i);
					    	Cell cell = row.getCell(Integer.parseInt(column.getText()));
					    	//System.out.println(cell.getStringCellValue());
					    	
					    	String parsed = null;
					    	try
					    	{
					    		parsed = cell.getStringCellValue();
					    	}
					    	
					    	catch(IllegalStateException e)
					    	{
					    		continue;
					    	}
					    	StringBuilder editedString = new StringBuilder(parsed);
						    int index = editedString.indexOf(delimiter.getText());
						    //System.out.println(index);
						    
						    if(index == -1)
						    {
						    	continue;
						    }
						    else
						    {
						    	String writeBack = editedString.substring(index + 1);
						    	cell.setCellValue(writeBack);
						    }
						    
					    	
					    	//cell.setCellValue("a test");
					    
				        //System.out.println("Filename: " + inputFilename.getText());
				        //System.out.println("column: " + column.getText());
				        //System.out.println("delimiter: " + delimiter.getText());
					    }
					    
					    FileOutputStream fileOut = new FileOutputStream(inputFilename.getText());
					    //FileOutputStream fileOut = new FileOutputStream("Carson Piping Cross Reference 10-9-2013 (Master).xls");
				    	wb.write(fileOut);
					    fileOut.close();
				      	
				    }
				}
				
				finally
				{
				
				}
				/*
				catch (FileNotFoundException e)
			    {System.out.println("File Not Found"); }
			    catch (IOException e)
			    {System.out.println("IO Error"); }
				*/
				
			/*
				InputStream inp = new FileInputStream(inputFileName);
				
			    Workbook wb = WorkbookFactory.create(inp);
			    org.apache.poi.ss.usermodel.Sheet sheet = wb.getSheetAt(0);
			    Row row = ((org.apache.poi.ss.usermodel.Sheet) sheet).getRow(0);
			    Cell cell = row.getCell(7);
			    //if (cell == null)
			       // cell = row.createCell(3);
			    //cell.setCellType(Cell.CELL_TYPE_STRING);
			    cell.setCellValue("a test");
			    
	    
		    // Write the output to a file
	
		    	FileOutputStream fileOut = new FileOutputStream("Carson Piping Cross Reference 10-9-2013 (Master).xls");
		    	wb.write(fileOut);
			    fileOut.close();
				*/
				
				
			}
		    
			/*
			JTextFieldtextboxPanel xField = new JTextField(5);
		      JTextField yField = new JTextField(5);

		      JPanel myPanel = new JPanel();
		      myPanel.add(new JLabel("x:"));
		      myPanel.add(xField);
		      myPanel.add(Box.createHorizontalStrut(15)); // a spacer
		      myPanel.add(new JLabel("y:"));
		      myPanel.add(yField);

		      int result = JOptionPane.showConfirmDialog(null, myPanel, 
		               "Please Enter X and Y Values", JOptionPane.OK_CANCEL_OPTION);
		      if (result == JOptionPane.OK_OPTION) {
		         System.out.println("x value: " + xField.getText());
		         System.out.println("y value: " + yField.getText());
		      }	
			*/		
	   
	    		
		System.out.println("Finished Editing");

	}
	


	@Override
	public void actionPerformed(ActionEvent arg0) 
	{
		// This is where I do the chooseFile dialog
		
		JFileChooser chooseFile = new JFileChooser();
		
	   // JFileChooser chooser = new JFileChooser();
	    FileNameExtensionFilter filter = new FileNameExtensionFilter(
	        "Excel Spreadsheets", "xlsx", "xml", "xls");
	    chooseFile.setFileFilter(filter);
	    int returnVal = chooseFile.showOpenDialog(null);
	    if(returnVal == JFileChooser.APPROVE_OPTION) 
	    {
	    	StringBuilder tempString = new StringBuilder(chooseFile.getSelectedFile().getName());
	    	int tempIndex = tempString.indexOf(".");
	    	String extension = tempString.substring (tempIndex);
	    	if(extension.equals(".xls") || extension.equals(".xml") || extension.equals(".xlsx"))
	    	{
	    	inputFilename.setText(chooseFile.getSelectedFile().getName());
	    	}
	    	else
	    	{
	    		Object[] options = {"OK"};
	    	    int n = JOptionPane.showOptionDialog(null,
	    	    		"Invalid File Selection. Please select a .xls, .xml, or .xlsx.","Invalid Selection",
	    	                   JOptionPane.PLAIN_MESSAGE,
	    	                   JOptionPane.QUESTION_MESSAGE,
	    	                   null,
	    	                   options,
	    	                   options[0]);
	    		//JOptionPane.showConfirmDialog(null, "Invalid File Selection. Please select a .xls, .xml, or .xlsx.");
	    	}
	    } 	
	}

}
