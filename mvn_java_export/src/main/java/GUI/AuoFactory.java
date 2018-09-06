package GUI;

import java.awt.Component;
import java.awt.EventQueue;
import java.awt.Font;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Vector;

import javax.swing.DefaultCellEditor;
import javax.swing.JComboBox;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.WindowConstants;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableCellRenderer;
import javax.swing.table.TableColumn;

import OutputMethod.OutputCommon;

import javax.swing.JTable;
import javax.swing.SwingConstants;
import javax.swing.JScrollPane;
import javax.swing.JButton;
import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;

public class AuoFactory {

	private JFrame frame;
	private JTable table;
	private JButton btnNewButton;
	private JButton btnAdd;
	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					AuoFactory window = new AuoFactory();
					window.frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public AuoFactory() {
		initialize();
		
		DefaultTableModel tableModel = new DefaultTableModel();
		try {
			tableModel = buildTableModel(getData());
			table.setModel(tableModel);
			
//			table.getColumnModel().getColumn(0).setCellEditor(new DefaultCellEditor(generateBox()));
			
			DefaultTableCellRenderer centerRenderer = new DefaultTableCellRenderer();
			centerRenderer.setHorizontalAlignment(SwingConstants.CENTER);
			table.setDefaultRenderer(String.class, centerRenderer);
			
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		frame.setVisible(true);
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frame = new JFrame();
		frame.setBounds(100, 100, 529, 423);
		frame.setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);
		frame.getContentPane().setLayout(null);
		
		table = new JTable() {
			@Override
			public Component prepareRenderer(TableCellRenderer renderer, int row, int column) {
				Component component = super.prepareRenderer(renderer, row, column);
				int rendererWidth = component.getPreferredSize().width;
				TableColumn tableColumn = getColumnModel().getColumn(column);
				tableColumn.setPreferredWidth(Math.max(rendererWidth + getIntercellSpacing().width, tableColumn.getPreferredWidth()));
				
				return component;
			}
		};
		table.setRowSelectionAllowed(false);
		table.setFont(new Font("新細明體", Font.PLAIN, 17));
		table.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
		table.setRowHeight(20);
	
		JScrollPane scrollPane = new JScrollPane(table);
		scrollPane.setSize(400, 365);
		scrollPane.setLocation(10, 10);
		frame.getContentPane().add(scrollPane);
		
		// Save function
		btnNewButton = new JButton("Save");
		btnNewButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				save();
			}
		});
		btnNewButton.setBounds(420, 10, 87, 23);
		frame.getContentPane().add(btnNewButton);
	
		btnAdd = new JButton("Add");
		btnAdd.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				addNewRow();
			}
		});
		btnAdd.setBounds(420, 43, 87, 23);
		frame.getContentPane().add(btnAdd);

	}
	
	private void addNewRow() {
		DefaultTableModel model = (DefaultTableModel) table.getModel();
		model.addRow(new Object[]{" ", " "});
		table.getColumnModel().getColumn(0).setCellEditor(new DefaultCellEditor(generateBox()));		
	}
	
	private void save() {

		for (int i = 0; i < table.getRowCount(); i++) {
			String code=table.getValueAt(i, 0).toString();
			String factory=table.getValueAt(i, 1).toString();
			System.out.println(code+"="+factory);
			if (code==null || code.trim().length()<=0 || factory==null || factory.trim().length()<=0) {
				JOptionPane.showMessageDialog(null, "需都有值!!!", "錯誤", JOptionPane.ERROR_MESSAGE);
			} else {
				if (!saveToDB(code,factory)) {
					JOptionPane.showMessageDialog(null,"保存失敗!!!", "錯誤",JOptionPane.ERROR_MESSAGE);
				}
			}
		}
		
	}
	
	private ResultSet getData() throws Exception {
		Connection conn = OutputCommon.connSQL();
		
		String sql = " select CODE,AUO_FACTORY from D_CUST_AUO ";
		Statement stat = conn.createStatement();
		return stat.executeQuery(sql);
	}
	
	private JComboBox generateBox() {
		JComboBox bx = new JComboBox();
		try {
			Connection con = OutputCommon.connSQL();
			String sql = " select CODE from D_CUST where ID='16130599' ";
			Statement stat = con.createStatement();
			ResultSet rs = stat.executeQuery(sql);
			while(rs.next()){
				bx.addItem(rs.getString(1));
	        }
		} catch (Exception e) {
			e.printStackTrace();
		}
	     
	    return bx;
	}
	
	private boolean saveToDB(String code, String factory) {
		boolean result = false;
		
		try {
			Connection conn = OutputCommon.connSQL();
			String sql = "  select * from  D_CUST_AUO where COMP_ID = 'A' and CODE=ltrim('"+code+"')";
			Statement stat = conn.createStatement();
			
			ResultSet rs = stat.executeQuery(sql);
			if (rs.next()) {
				sql = " UPDATE D_CUST_AUO set AUO_FACTORY='"+factory.trim()+"' where COMP_ID='A' and CODE=ltrim('"+code+"') ";
			} else {
				sql = " INSERT INTO D_CUST_AUO values ('A','"+code.trim()+"','"+factory.trim()+"')";
			}
			System.out.println(sql);
			int cnt = stat.executeUpdate(sql);
			if (cnt>0) {
				result = true;
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		return result;
	}

	/**
	 * 將ResultSet lay 進JTable 裡
	 * 
	 * @param rs
	 * @return
	 */
	public static DefaultTableModel buildTableModel(ResultSet rs) {
		// names of columns
		Vector<String> columnNames;
		// data of the table
		Vector<Vector<Object>> data;
		try {
			ResultSetMetaData metaData = rs.getMetaData();

			columnNames = new Vector<String>();
			int columnCount = metaData.getColumnCount();
			
			for (int column = 1; column <= columnCount; column++) {
				columnNames.add(metaData.getColumnName(column));
			}

			data = new Vector<Vector<Object>>();
			while (rs.next()) {
				Vector<Object> vector = new Vector<Object>();
				for (int columnIndex = 1; columnIndex <= columnCount; columnIndex++) {
					vector.add(rs.getObject(columnIndex));
				}
				data.add(vector);
			}
			return new DefaultTableModel(data, columnNames);
		} catch (SQLException e) {
			e.printStackTrace();
		}

		return new DefaultTableModel();
	}
	
}
