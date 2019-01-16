// Coded by Travis Kehrli
// Last updated 2/7/2018
// The purpose of this program is to take all text files in the given directory
//   and output them into an Excel spreadsheet such that one column is the name
//   of the file and the second column is the contents of the file.
//  It can be run from inside the target folder or a path to the target folder may be given.
//  Of course there are some limitations, such as max rows/columns and max characters per cell.

package textfilestoexcel;

import java.io.File;
import java.io.FileOutputStream;
import java.nio.file.Files;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class TextFiles_to_Excel {
	
	public static void main(String[] args) {

		try {
			Workbook wb = new HSSFWorkbook(); //make a workbook
			CreationHelper createHelper = wb.getCreationHelper(); //create a helper
			Sheet sheet1 = wb.createSheet("Sheet 1"); //create a sheet
			Row row; //get the variable ready
			
			row = sheet1.createRow((short)0); //make the first row
		    row.createCell(1).setCellValue(createHelper.createRichTextString("File Names")); //set the column #1 name
		    row.createCell(2).setCellValue(createHelper.createRichTextString("File Contents")); //set the column #2 name
			
		    File dir; //get the variable ready
			if(args.length != 0) { //if they did give us an arguement
				dir = new File(args[0]); //then use it
				
			} else { //otherwise
				dir = new File(System.getProperty("user.dir")); //just use the current directory
				//dir = new File("/Users/main/Desktop/export_target");
			}
			
			
			
			File[] directoryListing = dir.listFiles(); //get the directory listing
			int rowCount = 1; //initiate the count
			int columnCount = 0;
			int totalCount = 0;
			
			
			if (directoryListing != null) { //if it is a directory
				for (File child : directoryListing) { //for each file in the directory

					
					String extension = "";

					int i = child.getAbsolutePath().lastIndexOf('.');
					if (i > 0) {
					    extension = child.getAbsolutePath().substring(i+1);
					}
					
					if(extension.equals("txt")) {
						String fileName = child.getName(); //get the name of the file
						String fileContents = new String(Files.readAllBytes(child.toPath()));; //get the contents of the file
						
						
						
						row = sheet1.createRow((short)rowCount); //make a new row based on the count
					    row.createCell(columnCount).setCellValue(createHelper.createRichTextString(fileName)); //set the first cell as the name
					    row.createCell(columnCount+1).setCellValue(fileContents.toString()); //set the second cell as the contents
					    rowCount++; //increment the count
					    totalCount++;
					    
					    if(rowCount%5000 == 0) {
					    		System.out.println("On row number " + rowCount);
					    }
					    
					    if(rowCount >= 32768) {
					    		columnCount = columnCount + 3;
					    		System.out.println("Hit row " + rowCount + ", starting new pair of columns at position " + columnCount);
					    		rowCount = 1;
					    }
					}
						
					
				}
				
				System.out.println("Processing a total of " + totalCount + " text files. Please wait.");
				FileOutputStream fileOut = new FileOutputStream("AllTxtFiles.xls"); //make a new file output stream
			    wb.write(fileOut); //write it
			    fileOut.close(); //close it
			    wb.close();
			    System.out.println("Finished processing. XLS file created. Have a good day.");
				
			} else { //if it is not a directory
				//tell them
				System.out.println("That is not a directory.");
				System.out.println("Proper usage is  txtToXL <folder containing txt files>");
				System.out.println("If no directory is specified, the current working directory will be used.");
			}
			
		} catch (Exception e) {
			System.out.println("Error: ");
			e.printStackTrace();
		}
		
	}
}
