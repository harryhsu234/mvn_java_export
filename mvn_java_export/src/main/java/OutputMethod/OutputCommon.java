package OutputMethod;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Hashtable;
import java.util.Iterator;

import javax.swing.JOptionPane;
import javax.swing.filechooser.FileFilter;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class OutputCommon {
	
	String outputFilePath = "D:\\XML_OUTPUT\\";
	String outputFileName = "BENQ_"+System.currentTimeMillis()+".xlsx";
	
	public OutputCommon() {
	
	}
	
	public static Connection connSQL() throws Exception {
		String userName = "thi_mis";
		String password = "tecthi8686";
		String server = "192.168.0.222";

		//Date sys_date = new Date();
		String dbName = "";
		//SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");

		// if(sys_date.before( sdf.parse("2017/04/01") ))
		// dbName = "GICHIQDB";
		// else
		dbName = "GICWEBDB";

		String url = "jdbc:sqlserver://" + server + ";databaseName=" + dbName + "";

		Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
		Connection conn = DriverManager.getConnection(url, userName, password);

		return conn;
	}
	
	public static final String PREFIX = "stream2file";
    public static final String SUFFIX = ".tmp";

    public static File stream2file (InputStream in) throws IOException {
        final File tempFile = File.createTempFile(PREFIX, SUFFIX);
        tempFile.deleteOnExit();
        try (FileOutputStream out = new FileOutputStream(tempFile)) {
            IOUtils.copy(in, out);
        }
        return tempFile;
    }
	
	public static void infoBox(String infoMessage, String titleBar) {
		JOptionPane.showMessageDialog(null, infoMessage, titleBar, JOptionPane.INFORMATION_MESSAGE);
	}
	
	@SuppressWarnings("deprecation")
	public static String xls2xlsx(String xls_path, String xlsx_path) throws IOException {
		InputStream in = new BufferedInputStream(new FileInputStream(xls_path));
        try {
            Workbook wbIn = new HSSFWorkbook(in);
            File outF = new File(xlsx_path);
            if (outF.exists())
                outF.delete();

            Workbook wbOut = new XSSFWorkbook();
            int sheetCnt = wbIn.getNumberOfSheets();
            for (int i = 0; i < sheetCnt; i++) {
                Sheet sIn = wbIn.getSheetAt(i);
                Sheet sOut = wbOut.createSheet(sIn.getSheetName());
                Iterator<Row> rowIt = sIn.rowIterator();
                while (rowIt.hasNext()) {
                    Row rowIn = rowIt.next();
                    Row rowOut = sOut.createRow(rowIn.getRowNum());

                    Iterator<Cell> cellIt = rowIn.cellIterator();
                    while (cellIt.hasNext()) {
                        Cell cellIn = cellIt.next();
                        Cell cellOut = rowOut.createCell(
                                cellIn.getColumnIndex(), cellIn.getCellType());

                        switch (cellIn.getCellType()) {
                        case Cell.CELL_TYPE_BLANK:
                            break;

                        case Cell.CELL_TYPE_BOOLEAN:
                            cellOut.setCellValue(cellIn.getBooleanCellValue());
                            break;

                        case Cell.CELL_TYPE_ERROR:
                            cellOut.setCellValue(cellIn.getErrorCellValue());
                            break;

                        case Cell.CELL_TYPE_FORMULA:
//                        	
                        	String f = cellIn.getCellFormula();
                        	System.out.println(sIn.getSheetName() + " : " +cellIn.getAddress());
                        	try {
                        		System.out.println("Formula is " + cellIn.getCellFormula());
                                switch(cellIn.getCachedFormulaResultType()) {
                                    case Cell.CELL_TYPE_NUMERIC:
                                        System.out.println("Last evaluated as: " +cellIn.getNumericCellValue());
                                        cellOut.setCellType(Cell.CELL_TYPE_NUMERIC);
                                        cellOut.setCellValue(cellIn.getNumericCellValue());
                                        break;
                                    case Cell.CELL_TYPE_STRING:
                                        System.out.println("Last evaluated as \"" + cellIn.getRichStringCellValue() + "\"");
                                        cellOut.setCellType(Cell.CELL_TYPE_STRING);
                                        cellOut.setCellValue(cellIn.getRichStringCellValue().getString());
                                        break;
                                }
                            
                            }
                        	catch(Exception ex) {
                        		System.err.print("忽略FORMULA: ");
                        		System.err.println(f);
                        		System.err.println(ex.getMessage());
                        	}
                            break;

                        case Cell.CELL_TYPE_NUMERIC:
                            cellOut.setCellValue(cellIn.getNumericCellValue());
                            break;

                        case Cell.CELL_TYPE_STRING:
                            cellOut.setCellValue(cellIn.getStringCellValue());
                            break;
                        }

                        {
                            CellStyle styleIn = cellIn.getCellStyle();
                            CellStyle styleOut = cellOut.getCellStyle();
                           	styleOut.setDataFormat(styleIn.getDataFormat());
                        }
                        cellOut.setCellComment(cellIn.getCellComment());

                        // HSSFCellStyle cannot be cast to XSSFCellStyle
                        // cellOut.setCellStyle(cellIn.getCellStyle());
                    }
                }
            }
            OutputStream out = new BufferedOutputStream(new FileOutputStream(xlsx_path));
            try {
                wbOut.write(out);
            } finally {
                out.close();
            }
        } finally {
            in.close();
        }
		
		return xlsx_path;
	}
	
/***
 * 將值寫入WORKSHEET
 * @param ws 寫入WORKSHEET 目標
 * @param row_pos 第幾個ROW (1-BASE)
 * @param col_pos 第幾個COLUMN (1-BASE)
 * @param value
 * @throws Exception
 */
	protected void setValue(XSSFSheet ws, int row_pos, int col_pos, Object value) throws Exception {
		// create and set cell
		if (ws.getRow(row_pos) == null) {
			ws.createRow(row_pos);
		}
		if (ws.getRow(row_pos).getCell(col_pos) == null) {
			XSSFCell newCell = ws.getRow(row_pos).createCell(col_pos);
		}

		// 判斷填入值是否為NULL
		if (value == null) {
			System.err.println("Value is null");
			return;
		}
		String className = value.getClass().getName();
		if (className == "java.lang.Integer")
			ws.getRow(row_pos).getCell(col_pos).setCellValue((Integer) value);
		else if (className == "java.lang.Double")
			ws.getRow(row_pos).getCell(col_pos).setCellValue((Double) value);
		else if (className == "java.lang.String")
			ws.getRow(row_pos).getCell(col_pos).setCellValue((String) value);
		else if(className == "java.sql.Timestamp")
			ws.getRow(row_pos).getCell(col_pos).setCellValue(new SimpleDateFormat("MM/dd/yyyy").format(value));
		else
			throw new Exception("Cell format not supported: " + className);
	}
	
	/***
	 * 將值寫入WORKSHEET
	 * @param ws 寫入WORKSHEET 目標
	 * @param row_pos 第幾個ROW (1-BASE)
	 * @param col_pos 第幾個COLUMN (1-BASE)
	 * @param value
	 * @throws Exception
	 */
		protected void setValue(HSSFSheet ws, int row_pos, int col_pos, Object value) throws Exception {
			// create and set cell
			if (ws.getRow(row_pos) == null) {
				ws.createRow(row_pos);
			}
			if (ws.getRow(row_pos).getCell(col_pos) == null) {
				HSSFCell newCell = ws.getRow(row_pos).createCell(col_pos);
			}

			// 判斷填入值是否為NULL
			if (value == null) {
				System.err.println("Value is null");
				return;
			}
			String className = value.getClass().getName();
			if (className == "java.lang.Integer")
				ws.getRow(row_pos).getCell(col_pos).setCellValue((Integer) value);
			else if (className == "java.lang.Double")
				ws.getRow(row_pos).getCell(col_pos).setCellValue((Double) value);
			else if (className == "java.lang.String")
				ws.getRow(row_pos).getCell(col_pos).setCellValue((String) value);
			else if(className == "java.sql.Timestamp")
				ws.getRow(row_pos).getCell(col_pos).setCellValue(new SimpleDateFormat("MM/dd/yyyy").format(value));
			else
				throw new Exception("Cell format not supported: " + className);
		}
		
	
	/***
	 * 抓第 i 個Cell 的String 值 ( 0-base )
	 * 
	 * @param i
	 * @return
	 */
	protected String getCellString(Row row, int i) {
		String result = "";
		try {
			result = String.valueOf(row.getCell(i).getNumericCellValue());
			if (result.equals("0")||result.equals("0.0")) {
				result = "";
			}
		} catch (Exception ex) {
			String exm = ex.getMessage();
//			ex.printStackTrace();
		}
		try {
			result = row.getCell(i).getStringCellValue();
		} catch (Exception ex) {
			String exm = ex.getMessage();
//			ex.printStackTrace();
		}
		
		return result;
	}

	/***
	 * 抓第 i 個Cell 的double 值 ( 0-base )
	 * 
	 * @param i
	 * @return
	 */
	protected double getCellDouble(Row row, int i) {
		try {
			return row.getCell(i).getNumericCellValue();
		} catch (Exception ex) {
			String exm = ex.getMessage();
			return 0;
		}
	}

	protected class ExcelFilter extends FileFilter {
		@Override
		public boolean accept(File pathname) {
			String filename = pathname.getName();
			if (pathname.isDirectory()) {
				return true;

			} else if (filename.endsWith("xls") || filename.endsWith("xlsx")) {
				return true;
			} else {
				return false;
			}
		}

		@Override
		public String getDescription() {
			return "Excel Files";
		}
	}
	
	/***
	 * 將ResultSet 轉成ArrayList<Hashtable>
	 * @param rs
	 * @return
	 * @throws SQLException 
	 */
	protected ArrayList<Hashtable<String,Object>> RS2AL(ResultSet rs) throws SQLException {
		ArrayList<Hashtable<String,Object>> arraylist = new ArrayList<Hashtable<String,Object>>();
		ResultSetMetaData meta = rs.getMetaData();
		
		int metaColCount = meta.getColumnCount();
		while(rs.next()) {
			Hashtable<String, Object> ht = new Hashtable<String, Object>();
			
			for(int i = 1; i <= metaColCount; i++) {
				String col_name = meta.getColumnLabel(i);
				String col_type = meta.getColumnClassName(i);
				
				Object value = null;
				if (col_type == "java.lang.Integer") 
					value = rs.getInt(i);
				else if (col_type == "java.lang.Double")
					value = rs.getDouble(i);
				else if (col_type == "java.lang.String") 
					value = rs.getString(i);
				else if(col_type == "java.sql.Timestamp")
					value = rs.getTimestamp(i);
				else
					infoBox("[CALL JASON_THITWN#174] DATA TYPE "+col_type+" not spported", "CALL THITWN#174");
				
				if(rs.wasNull()) {
					if (col_type == "java.lang.Integer") 
						value = 0;
					else if (col_type == "java.lang.Double")
						value = 0;
					else if (col_type == "java.lang.String") 
						value = "";
					else if(col_type == "java.sql.Timestamp")
						value = new Timestamp(0);
				}
				
				ht.put(col_name, value);
			}
			
			arraylist.add(ht);
		}
		
		rs.first();
		
		return arraylist;
	}
}
