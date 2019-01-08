package search.excel;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SearchExcel {
	SearchExcel(){}

	public List<String> listFilesForFolder(final File folder){
		List<String> res = new ArrayList<>();
		
		for (final File fileEntry : folder.listFiles()) {
	        if (fileEntry.isDirectory()) {
	            listFilesForFolder(fileEntry);
	        } else {
	        	res.add(folder.toString().replace("\\", "/")+"/"+fileEntry.getName());
	        }
	    }
		return res;
	}
	public List<String[]> getexcel(String file_path, String[] input_array) {
			/*String[] arr = file_path.split("/");
			String[] arr2 = arr[arr.length-1].split("\\.");
			
			String filename = arr2[0];*/
		List<String[]> res = new ArrayList<>();
			try {
				OPCPackage pkg = OPCPackage.open(new File(file_path));
			   // POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(file));
			    XSSFWorkbook wb = new XSSFWorkbook(pkg);
			    for(int sheet_index = 0; sheet_index<wb.getNumberOfSheets();sheet_index++) {
			    XSSFSheet sheet = wb.getSheetAt(sheet_index);
			    XSSFRow row;
			    XSSFCell cell;
			    int rows; // No of rows
			    rows = sheet.getPhysicalNumberOfRows();
		
			    int cols = 0; // No of columns
			    int tmp = 0;
			    
			    // This trick ensures that we get the data properly even if it doesn't start from first few rows
			    for(int i = 0; i < 20 || i < rows; i++) {
			        row = sheet.getRow(i);
			        if(row != null) {
			            tmp = sheet.getRow(i).getPhysicalNumberOfCells();
			            if(tmp > cols) cols = tmp;
			        }
			    }
		    	
		    	
			    for(int r = 0; r < rows; r++) {
			    	
			        row = sheet.getRow(r);
			        if(row != null) {
			        	String[] each_row = new String[cols+1];
			            for(int c = 0; c < cols; c++) {
			                cell = row.getCell(c);
			                if(cell != null) {
			                	
			                  if(isInArray(cell.toString(), input_array)) {
			                	  each_row[0]=file_path;
			                	  for(int i = 0; i<cols;i++) {
			                		  if(row.getCell(i)!=null) {
			                			  each_row[i+1]=row.getCell(i).toString();
			                			 
			                		  }
			                	  }
			                  }
			                }else {
			                	continue;
			                }
			            }
			            
			            if(each_row[0]!=null) {
			            res.add(each_row);
			            }
			        }
			    }
			    }
			} catch(Exception ioe) {
			    ioe.printStackTrace();
			}
		return res;
		
	}
	public boolean isInArray(String item,String [] array) {
		for(int i = 0; i<array.length; i++) {
			if(item.equals(array[i])) {
				return true;
			}
		}
		return false;
	}
		
		
		public void printExcel(List<String> allfilepaths, String[] arr) {
			SearchExcel e = new SearchExcel();
			List<String[]> list = new ArrayList<>();
			allfilepaths.forEach((item)->{
				list.addAll(e.getexcel(item, arr));
				System.out.println(item);
			});
			
		
			XSSFWorkbook workbook = new XSSFWorkbook();
		    XSSFSheet sheet = workbook.createSheet("All Done");
		     
		    int rowCount = 0;
		     
		    for (int i = 0 ; i < list.size(); i++) {
		        Row row = sheet.createRow(++rowCount);
		         
		        int columnCount = 0;
		         
		        for (int j = 0 ; j<list.get(i).length;j++) {
		        	if(j>16300) {
		        		continue;
		        	}
		            Cell cell = row.createCell(++columnCount);
		          
		                cell.setCellValue((String) list.get(i)[j]);
		           
		        }
		         
		    }
		    try (FileOutputStream outputStream = new FileOutputStream("All_Done"+".xlsx")) {
		        workbook.write(outputStream);
		    } catch (FileNotFoundException e1) {
				e1.printStackTrace();
			} catch (IOException e1) {
				e1.printStackTrace();
			}
	       //System.out.println("_Done"+".xlsx");
		}
		
		public List<String> getFileList(String parent_path){
			parent_path = parent_path.replace("\\", "/");
			File folder = new File(parent_path);
			List<String> filelist = new SearchExcel().listFilesForFolder(folder);
			return filelist;
		}
		/**
		 * @param args
		 * @throws IOException
		 */
		
}
