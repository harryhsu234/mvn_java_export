package GUI;

import java.awt.Color;
import java.awt.Component;
import java.awt.Dimension;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Properties;
import java.util.Vector;

import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JFormattedTextField.AbstractFormatter;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.ListSelectionModel;
import javax.swing.SwingConstants;
import javax.swing.WindowConstants;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableCellRenderer;
import javax.swing.table.TableColumn;

import org.jdatepicker.impl.JDatePanelImpl;
import org.jdatepicker.impl.JDatePickerImpl;
import org.jdatepicker.impl.UtilDateModel;

import OutputMethod.Excel_AE_BENQ;
import OutputMethod.Excel_AIDC;
import OutputMethod.Excel_C1_BENQ;
import OutputMethod.Excel_C2_BENQ;
import OutputMethod.Excel_C3_BENQ;
import OutputMethod.Excel_C3_CommonUse;
import OutputMethod.Excel_C4_BENQ;
import OutputMethod.Excel_C4_BENQ2;
import OutputMethod.Excel_C4_NOVARTIS;
import OutputMethod.OutputCommon;
import OutputMethod.Xml_AUO_Export;
import OutputMethod.Xml_AUO_Import;
import OutputMethod.Xml_WPG;
import javax.swing.JRadioButton;
import javax.swing.ButtonGroup;
import javax.swing.JComboBox;

public class FancyGUI {
	public static final int EXCEL_BENQ = 1001;// "EXCEL_BENQ";
	public static final int AE_EXCEL_BENQ = 1002;// "AE_EXCEL_BENQ";
	public static final int XML_WPG = 2000; // "XML_WPG";
	public static final int AE_XML_WPG = 2001; // "AE_XML_WPG";
	public static final int AI_XML_WPG = 2002; // "AI_XML_WPG";
	public static final int BR_XML_AUO_IMP = 3000; // "BR_XML_AUO_IMPORT";
	public static final int BR_XML_AUO_EXP = 3001; // "BR_XML_AUO_EXPORT";
	public static final int C1_EXCEL_BENQ = 4001; // "C1_EXCEL_BENQ"
	public static final int C2_EXCEL_BENQ = 4002; // "C2_EXCEL_BENQ"
	public static final int C3_EXCEL_BENQ = 4003; // "C2_EXCEL_BENQ"
	public static final int C4_EXCEL_BENQ = 4004; // "C4_EXCEL_BENQ"
	public static final int C4_EXCEL_BENQ2 = 4005; // "C4_EXCEL_BENQ"  C4_EXCEL_BENQ3
	public static final int C4_EXCEL_BENQ3 = 4006; // "C4_EXCEL_BENQ3"
	public static final int C3_EXCEL_COMMONUSE = 7001; // "C2_EXCEL_BENQ"
	public static final int C4_EXCEL_NOVARTIS = 7010; // "C4_EXCEL_BENQ"
	public static final int AIDC_EXCEL = 8001; // "AIDC_EXCEL"

	private String currentMkYear = "";
	private String isExpImp = ""; // AE or AI
	boolean isShowSelectAll = true;

	private JFrame frmFcyGui;
	private JTable table;
	private JTextField tfBooking;
	private JTextField tfBooking2;
	private JTextField tfMAWB;
	private JTextField tfHAWB;
	private JTextField tfYear;
	private JTextField tfCartonno;
	private JTextField tfRFive;
	private JTextField tfRFive2;
	private JLabel label_3;

	private JDatePickerImpl dpDrDate;
	private JDatePickerImpl dpDrDate2;
	private JLabel label_4;

	private JButton selectAllBtn;
	private JButton searchExpBtn;
	private JButton searchImpBtn;
	private JButton setParameter;
	private JButton btnAuo;
	private JTextField tfSHPR;
	private JCheckBox bRelease;
	private int assignedProg = XML_WPG;
	private String configPrefix = "";
	private final ButtonGroup buttonGroup = new ButtonGroup();
	
	JRadioButton rdbtnAIR = new JRadioButton("空運報單");
	JRadioButton rdbtnSEA = new JRadioButton("海運報單");
	private JLabel label_6;
	JComboBox chooser;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					FancyGUI window = new FancyGUI();
					window.frmFcyGui.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public FancyGUI() {
		initialize();
	}

	public FancyGUI(int aeXmlWpg) {
		initialize();
		this.assignedProg = aeXmlWpg;

		switch (aeXmlWpg) {
		case AE_XML_WPG:
			frmFcyGui.setTitle(Xml_WPG.programmeTitle);
			searchImpBtn.setEnabled(false);
			rdbtnAIR.setSelected(true);
			break;
		case AE_EXCEL_BENQ:
			frmFcyGui.setTitle(Excel_AE_BENQ.programmeTitle);
			searchImpBtn.setEnabled(false);
			rdbtnAIR.setSelected(true);
			break;

		case AI_XML_WPG:
			frmFcyGui.setTitle(Xml_WPG.programmeTitle);
			searchExpBtn.setEnabled(false);
			rdbtnAIR.setSelected(true);
			break;

		case XML_WPG:
			frmFcyGui.setTitle(Xml_WPG.programmeTitle);
			break;
		case EXCEL_BENQ:
			frmFcyGui.setTitle(Excel_AE_BENQ.programmeTitle);
			rdbtnAIR.setSelected(true);
			break;
		case BR_XML_AUO_IMP:
			frmFcyGui.setTitle(Xml_AUO_Import.programmeTitle);
			configPrefix = "GLS_IMPORT";
			setParameter.setEnabled(true);
			btnAuo.setEnabled(true);
			break;
		case BR_XML_AUO_EXP:
			frmFcyGui.setTitle(Xml_AUO_Export.programmeTitle);
			configPrefix = "GLS_EXPORT";
			setParameter.setEnabled(true);
			btnAuo.setEnabled(true);
			break;
		case C1_EXCEL_BENQ:
			frmFcyGui.setTitle(Excel_C1_BENQ.programmeTitle);
			searchImpBtn.setEnabled(false);
			rdbtnSEA.setSelected(true);
			break;
		case C2_EXCEL_BENQ:
			frmFcyGui.setTitle(Excel_C2_BENQ.programmeTitle);
			searchExpBtn.setEnabled(false);
			// rdbtnSEA.setSelected(true);
			break;
		case C3_EXCEL_BENQ:
			frmFcyGui.setTitle(Excel_C3_BENQ.programmeTitle);
			searchImpBtn.setEnabled(false);
			rdbtnAIR.setSelected(true);
			break;
		case C4_EXCEL_BENQ:
			frmFcyGui.setTitle(Excel_C4_BENQ.programmeTitle);
			searchExpBtn.setEnabled(false);
			rdbtnAIR.setSelected(true);
			break;
		case C4_EXCEL_BENQ2:
			frmFcyGui.setTitle(Excel_C4_BENQ2.programmeTitle);
			searchExpBtn.setEnabled(false);
			rdbtnAIR.setSelected(true);
			break;
		case C4_EXCEL_BENQ3:
			frmFcyGui.setTitle(Excel_C2_BENQ.programmeTitle);
			searchExpBtn.setEnabled(false);
			rdbtnAIR.setSelected(true);
			break;
		case C3_EXCEL_COMMONUSE:
			frmFcyGui.setTitle(Excel_C3_BENQ.programmeTitle);
			searchImpBtn.setEnabled(false);
			rdbtnAIR.setSelected(true);
			break;
		case C4_EXCEL_NOVARTIS:
			frmFcyGui.setTitle(Excel_C4_NOVARTIS.programmeTitle);
			searchExpBtn.setEnabled(false);
			rdbtnAIR.setSelected(true);
			break;
		case AIDC_EXCEL:
			frmFcyGui.setTitle(Excel_AIDC.programmeTitle);
//			searchExpBtn.setEnabled(false);
			rdbtnAIR.setSelected(true);
			break;
		default:
			JOptionPane.showMessageDialog(null, "sub-Program 選擇器回傳問題", "錯誤 GUI-002", JOptionPane.INFORMATION_MESSAGE);
			break;
		}
		

		this.frmFcyGui.setVisible(true);
		this.tfRFive.requestFocusInWindow();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	@SuppressWarnings("serial")
	private void initialize() {
		// 22 is harry good friend.
		frmFcyGui = new JFrame();
		frmFcyGui.setTitle("title set by program, if you see this, then there's problem");
		frmFcyGui.setBounds(100, 100, 806, 631);
		frmFcyGui.setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);
		frmFcyGui.getContentPane().setLayout(null);

		Dimension dim = Toolkit.getDefaultToolkit().getScreenSize();
		frmFcyGui.setLocation(dim.width / 2 - frmFcyGui.getSize().width / 2,
				dim.height / 2 - frmFcyGui.getSize().height / 2);

		JButton runBtn = new JButton("\u8F49\u51FA");
		runBtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				runXML();
			}
		});
		runBtn.setBounds(669, 125, 99, 27);
		frmFcyGui.getContentPane().add(runBtn);

		searchExpBtn = new JButton("出口報單查詢");
		searchExpBtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				isExpImp = "EXP";
				search();
			}
		});
		searchExpBtn.setBounds(384, 125, 128, 27);
		frmFcyGui.getContentPane().add(searchExpBtn);

		JScrollPane scrollPane = new JScrollPane();
		scrollPane.setBounds(14, 165, 754, 375);
		frmFcyGui.getContentPane().add(scrollPane);

		table = new JTable() {
			@Override
			public Class<?> getColumnClass(int column) {
				switch (column) {
				case 0:
					return Boolean.class;
				default:
					return String.class;
				}
			}

			@Override
			public Component prepareRenderer(TableCellRenderer renderer, int row, int column) {
				Component component = super.prepareRenderer(renderer, row, column);
				int rendererWidth = component.getPreferredSize().width;
				TableColumn tableColumn = getColumnModel().getColumn(column);
				tableColumn.setPreferredWidth(
						Math.max(rendererWidth + getIntercellSpacing().width, tableColumn.getPreferredWidth()));
				return component;
			}
		};
		table.setRowSelectionAllowed(false);
		table.setFont(new Font("新細明體", Font.PLAIN, 17));
		table.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
		table.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);
		table.setRowHeight(20);

		scrollPane.setViewportView(table);

		JLabel lblNewLabel = new JLabel("\u6587\u4EF6\u7DE8\u865F");
		lblNewLabel.setFont(new Font("新細明體", Font.PLAIN, 18));
		lblNewLabel.setBounds(14, 14, 80, 19);
		frmFcyGui.getContentPane().add(lblNewLabel);

		tfBooking = new JTextField();
		tfBooking.setFont(new Font("新細明體", Font.PLAIN, 18));
		tfBooking.setBounds(90, 11, 116, 25);
		frmFcyGui.getContentPane().add(tfBooking);
		tfBooking.setColumns(10);

		JLabel label = new JLabel("~");
		label.setFont(new Font("新細明體", Font.PLAIN, 18));
		label.setBounds(208, 14, 20, 19);
		frmFcyGui.getContentPane().add(label);

		tfBooking2 = new JTextField();
		tfBooking2.setFont(new Font("新細明體", Font.PLAIN, 18));
		tfBooking2.setColumns(10);
		tfBooking2.setBounds(220, 11, 116, 25);
		frmFcyGui.getContentPane().add(tfBooking2);

		JLabel lblMawb = new JLabel("MAWB");
		lblMawb.setFont(new Font("新細明體", Font.PLAIN, 18));
		lblMawb.setBounds(372, 14, 75, 19);
		frmFcyGui.getContentPane().add(lblMawb);

		tfMAWB = new JTextField();
		tfMAWB.setFont(new Font("新細明體", Font.PLAIN, 18));
		tfMAWB.setColumns(10);
		tfMAWB.setBounds(450, 11, 116, 25);
		frmFcyGui.getContentPane().add(tfMAWB);

		JLabel lblHawb = new JLabel("HAWB");
		lblHawb.setFont(new Font("新細明體", Font.PLAIN, 18));
		lblHawb.setBounds(580, 14, 75, 19);
		frmFcyGui.getContentPane().add(lblHawb);

		tfHAWB = new JTextField();
		tfHAWB.setFont(new Font("新細明體", Font.PLAIN, 18));
		tfHAWB.setColumns(10);
		tfHAWB.setBounds(652, 11, 116, 25);
		frmFcyGui.getContentPane().add(tfHAWB);

		JLabel label_1 = new JLabel("年度/箱號");
		label_1.setFont(new Font("新細明體", Font.PLAIN, 18));
		label_1.setBounds(14, 52, 87, 19);
		frmFcyGui.getContentPane().add(label_1);

		JLabel label_2 = new JLabel("~");
		label_2.setFont(new Font("新細明體", Font.PLAIN, 18));
		label_2.setBounds(348, 52, 20, 19);
		frmFcyGui.getContentPane().add(label_2);

		tfYear = new JTextField();
		tfYear.setFont(new Font("新細明體", Font.PLAIN, 18));
		currentMkYear = "" + (Calendar.getInstance().get(Calendar.YEAR) - 1911);
		currentMkYear = currentMkYear.substring(currentMkYear.length() - 2); // 設定為民國年末兩碼 2017 = 民國 106 產出 = 06
		tfYear.setText(currentMkYear);
		tfYear.setColumns(10);
		tfYear.setBounds(100, 49, 35, 25);
		frmFcyGui.getContentPane().add(tfYear);
		
		tfCartonno = new JTextField();
		tfCartonno.setFont(new Font("新細明體", Font.PLAIN, 18));
		tfCartonno.setColumns(10);
		tfCartonno.setBounds(147, 49, 58, 25);
		frmFcyGui.getContentPane().add(tfCartonno);

		tfRFive = new JTextField();
		tfRFive.setFont(new Font("新細明體", Font.PLAIN, 18));
		tfRFive.setColumns(10);
		tfRFive.setBounds(283, 49, 63, 25);
		frmFcyGui.getContentPane().add(tfRFive);

		tfRFive2 = new JTextField();
		tfRFive2.setFont(new Font("新細明體", Font.PLAIN, 18));
		tfRFive2.setColumns(10);
		tfRFive2.setBounds(358, 49, 63, 25);
		frmFcyGui.getContentPane().add(tfRFive2);

		label_3 = new JLabel("\u5831\u95DC\u65E5\u671F");
		label_3.setFont(new Font("新細明體", Font.PLAIN, 18));
		label_3.setBounds(14, 90, 80, 19);
		frmFcyGui.getContentPane().add(label_3);

		UtilDateModel model = new UtilDateModel();
		Properties p = new Properties();
		p.put("text.today", "Today");
		p.put("text.month", "Month");
		p.put("text.year", "Year");
		JDatePanelImpl datePanel = new JDatePanelImpl(model, p);
		dpDrDate = new JDatePickerImpl(datePanel, new LabelDateFormatter());
		dpDrDate.getJFormattedTextField().setFont(new Font("新細明體", Font.PLAIN, 18));
		dpDrDate.getJFormattedTextField().setBackground(Color.WHITE);
		dpDrDate.setSize(116, 27);
		dpDrDate.setLocation(95, 87);

		frmFcyGui.getContentPane().add(dpDrDate);

		UtilDateModel model2 = new UtilDateModel();
		JDatePanelImpl datePanel2 = new JDatePanelImpl(model2, p);
		dpDrDate2 = new JDatePickerImpl(datePanel2, new LabelDateFormatter());
		dpDrDate2.getJFormattedTextField().setFont(new Font("新細明體", Font.PLAIN, 18));
		dpDrDate2.getJFormattedTextField().setBackground(Color.WHITE);
		dpDrDate2.setSize(116, 27);
		dpDrDate2.setLocation(238, 87);

		frmFcyGui.getContentPane().add(dpDrDate2);

		label_4 = new JLabel("~");
		label_4.setFont(new Font("新細明體", Font.PLAIN, 18));
		label_4.setBounds(214, 90, 20, 19);
		frmFcyGui.getContentPane().add(label_4);

		tfSHPR = new JTextField();
		tfSHPR.setFont(new Font("新細明體", Font.PLAIN, 18));
		tfSHPR.setColumns(10);
		tfSHPR.setBounds(517, 87, 116, 25);
		frmFcyGui.getContentPane().add(tfSHPR);

		bRelease = new JCheckBox("\u5DF2\u653E\u884C");
		bRelease.setFont(new Font("新細明體", Font.PLAIN, 18));
		bRelease.setBounds(639, 87, 115, 27);
		frmFcyGui.getContentPane().add(bRelease);

		selectAllBtn = new JButton("全選");
		selectAllBtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				selectAll();
			}
		});
		selectAllBtn.setBounds(14, 125, 99, 27);
		frmFcyGui.getContentPane().add(selectAllBtn);

		JButton clearConBtn = new JButton("清除條件");
		clearConBtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				clearCondition();
			}
		});
		clearConBtn.setBounds(191, 125, 99, 27);
		frmFcyGui.getContentPane().add(clearConBtn);

		searchImpBtn = new JButton("進口報單查詢");
		searchImpBtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				isExpImp = "IMP";
				search();
			}
		});
		searchImpBtn.setBounds(527, 125, 128, 27);
		frmFcyGui.getContentPane().add(searchImpBtn);

		setParameter = new JButton("參數設定");
		setParameter.setEnabled(false);
		setParameter.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					new ParameterSetting(configPrefix);
				} catch (Exception e1) {
					e1.printStackTrace();
					JOptionPane.showMessageDialog(null, "自動錯誤訊息 - " + e1.getMessage(), "錯誤 GUI-015",
							JOptionPane.ERROR_MESSAGE);
				}
			}
		});
		setParameter.setBounds(14, 550, 87, 23);
		frmFcyGui.getContentPane().add(setParameter);

		btnAuo = new JButton("AUO 廠區");
		btnAuo.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					new AuoFactory();
				} catch (Exception e1) {
					e1.printStackTrace();
					JOptionPane.showMessageDialog(null, "自動錯誤訊息 - " + e1.getMessage(), "錯誤 GUI-016",
							JOptionPane.ERROR_MESSAGE);
				}
			}
		});
		btnAuo.setEnabled(false);
		btnAuo.setBounds(111, 550, 87, 23);
		frmFcyGui.getContentPane().add(btnAuo);
		
		rdbtnAIR.setFont(new Font("新細明體", Font.PLAIN, 18));
		buttonGroup.add(rdbtnAIR);
		rdbtnAIR.setBounds(536, 48, 113, 27);
		frmFcyGui.getContentPane().add(rdbtnAIR);
		
		rdbtnSEA.setFont(new Font("新細明體", Font.PLAIN, 18));
		buttonGroup.add(rdbtnSEA);
		rdbtnSEA.setBounds(655, 48, 113, 27);
		frmFcyGui.getContentPane().add(rdbtnSEA);
		
		JLabel label_5 = new JLabel("/");
		label_5.setFont(new Font("新細明體", Font.PLAIN, 18));
		label_5.setBounds(136, 55, 20, 19);
		frmFcyGui.getContentPane().add(label_5);
		
		label_6 = new JLabel("後五碼");
		label_6.setFont(new Font("新細明體", Font.PLAIN, 18));
		label_6.setBounds(218, 52, 87, 19);
		frmFcyGui.getContentPane().add(label_6);
		
		chooser = new JComboBox();
		chooser.setFont(new Font("新細明體", Font.PLAIN, 18));
		chooser.setBounds(384, 86, 128, 27);
		chooser.addItem("Shipper Code");
		chooser.addItem("Shipper Name");
		chooser.addItem("Consignee Code");
		chooser.addItem("Consignee Name");
		frmFcyGui.getContentPane().add(chooser);
	}

	public void selectAll() {

		for (int i = 0; i < table.getRowCount(); i++) {
			table.setValueAt(isShowSelectAll, i, 0);
		}

		isShowSelectAll = !isShowSelectAll;
		if (isShowSelectAll)
			selectAllBtn.setText("全選");
		else
			selectAllBtn.setText("清除全選");

	}

	public void clearCondition() {
		tfBooking.setText("");
		tfBooking2.setText("");
		tfMAWB.setText("");
		tfHAWB.setText("");
		tfYear.setText(currentMkYear);
		tfCartonno.setText("");
		tfRFive.setText("");
		tfRFive2.setText("");
		dpDrDate.getJFormattedTextField().setValue(null);
		dpDrDate2.getJFormattedTextField().setValue(null);

		tfSHPR.setText("");
		bRelease.setSelected(false);
	}

	public void search() {
		String sWhere = "";

		String sBooking = tfBooking.getText().trim();
		String sBooking2 = tfBooking2.getText().trim();
		if (!sBooking.isEmpty()) {
			if (!sBooking2.isEmpty()) {
				sWhere += " and DOC_HEAD_DOC_NO between '" + sBooking + "' and '" + sBooking2 + "' ";
			} else
				sWhere += " and DOC_HEAD_DOC_NO = '" + sBooking + "' ";
		}

		String sMAWB = tfMAWB.getText().trim();
		if (!sMAWB.isEmpty()) {
			sWhere += " and MAWB like '%" + sMAWB + "%' ";
		}

		String sHAWB = tfHAWB.getText().trim();
		if (!sHAWB.isEmpty()) {
			sWhere += " and HAWB like '%" + sHAWB + "%' ";
		}

		// tfYear.setText("");
		String sYear = tfYear.getText().trim();
		if (!sYear.isEmpty()) {
			sWhere += " and REPLACE(DCL_DOC_NO, ' ', '') like '%/" + sYear + "/%' ";
		}

		// tfCartonno
		String sCartonno = tfCartonno.getText().trim();
		if (!sCartonno.isEmpty()) {
			sWhere += " and REPLACE(DCL_DOC_NO, ' ', '') like '%/" + sCartonno + "/%' ";
		}
		
		String sRFive = tfRFive.getText().trim().toUpperCase();
		String sRFive2 = tfRFive2.getText().trim().toUpperCase();
		if (!sRFive.isEmpty()) {
			int sRFiveLength = sRFive.length();
			if (!sRFive2.isEmpty()) {
				int sRFive2Length = sRFive2.length();
				if (sRFiveLength != sRFive2Length) {
					JOptionPane.showMessageDialog(null, "報單號碼前後長度不同", "錯誤 GUI-001", JOptionPane.INFORMATION_MESSAGE);
					return;
				}

				sWhere += " and UPPER(RIGHT(RTRIM(DCL_DOC_NO)," + sRFiveLength + ")) between '" + sRFive + "' and '"
						+ sRFive2 + "' ";
			} else
				sWhere += " and UPPER(RIGHT(RTRIM(DCL_DOC_NO)," + sRFiveLength + ")) = '" + sRFive + "' ";
		}

		String sDrDt = dpDrDate.getJFormattedTextField().getText().trim();
		String sDrDt2 = dpDrDate2.getJFormattedTextField().getText().trim();
		if (!sDrDt.isEmpty()) {
			if (!sDrDt2.isEmpty()) {
				sWhere += " and DCL_DATE between '" + sDrDt + "' and '" + sDrDt2 + "' ";
			} else
				sWhere += " and DCL_DATE = '" + sDrDt + "' ";
		}

		String sSHPR = tfSHPR.getText().trim();
		if (!sSHPR.isEmpty()) {
			String sChooser = String.valueOf( chooser.getSelectedItem() );
			switch (sChooser) {
			case "Shipper Code":
				sWhere += " and UPPER(SHPR_CODE) like '%" + sSHPR.toUpperCase() + "%' ";
				break;
			case "Shipper Name":
				sWhere += " and UPPER(SHPR_E_NAME) like '%" + sSHPR.toUpperCase() + "%' ";
				break;
			case "Consignee Code":
				sWhere += " and UPPER(CNEE_CODE) like '%" + sSHPR.toUpperCase() + "%' ";
				break;
			case "Consignee Name":
				sWhere += " and UPPER(CNEE_E_NAME) like '%" + sSHPR.toUpperCase() + "%' ";
				break;
			}
		}
		
		/**
		if (!sSHPR.isEmpty()) {
			sWhere += " and UPPER(SHPR_CODE) like '%" + sSHPR.toUpperCase() + "%' ";
		}
		**/
		if (bRelease.isSelected()) {
			sWhere += " and RL_DATE is not null ";
		}
		
		// rdbtnSEA
		if (rdbtnAIR.isSelected()) {
			sWhere += " and AIR_SEA = '4' ";
		}
		if (rdbtnSEA.isSelected()) {
			sWhere += " and AIR_SEA = '1' ";
		}

		DefaultTableModel tableModel = new DefaultTableModel();

		try {
			tableModel = buildTableModel(getCustomInfo(sWhere, isExpImp));
		} catch (Exception e) {
			e.printStackTrace();
		}

		table.setModel(tableModel);

		// 文字置中顯示
		DefaultTableCellRenderer centerRenderer = new DefaultTableCellRenderer();
		centerRenderer.setHorizontalAlignment(SwingConstants.CENTER);
		table.setDefaultRenderer(String.class, centerRenderer);

	}

	/**
	 * @author Harry
	 * @return ArrayList<String> 畫面被選擇的報關單號
	 * @throws Exception selectedCustom.size() == 0 ,  "無報關單被選擇，請至少選一筆報關單"
	 */
	private ArrayList<String> getSelectedCustom() throws Exception {
		
		ArrayList<String> selectedCustom = new ArrayList<String>();

		for (int i = 0; i < table.getRowCount(); i++) {
			Boolean isChecked = Boolean.valueOf(table.getValueAt(i, 0).toString());
			String custom_no = table.getValueAt(i, 1).toString();
			String booking_no = table.getValueAt(i, 7).toString();
			if (isChecked) {
				selectedCustom.add(custom_no);
			}
		}

		if (selectedCustom.size() == 0)
			throw new Exception("無報關單被選擇，請至少選一筆報關單");

		return selectedCustom;
	}
	
	private ArrayList<String[]> getSelectedDrnoBKnoPair() throws Exception {
		ArrayList<String[]> selectedCustomPair = new ArrayList<String[]>();
		
		for (int i = 0; i < table.getRowCount(); i++) {
			Boolean isChecked = Boolean.valueOf(table.getValueAt(i, 0).toString());
			String custom_no = table.getValueAt(i, 1).toString();
			String booking_no = table.getValueAt(i, 7).toString();
			if (isChecked) {
				selectedCustomPair.add(new String[] { custom_no, booking_no });
			}
		}

		if (selectedCustomPair.size() == 0)
			throw new Exception("無報關單被選擇，請至少選一筆報關單");
		
		return selectedCustomPair;
	}
	

	public void runXML() {
		try {
			run_one_drno_to_one_output();
			
			run_many_drno_to_one_output();
		} // end try/catch

		catch (Exception e) {
			e.printStackTrace();
			JOptionPane.showMessageDialog(null, "自動錯誤訊息 - " + e.getMessage(), "錯誤 GUI-000", JOptionPane.ERROR_MESSAGE);
		}
	}
	
	private void run_many_drno_to_one_output() throws Exception {
		System.out.println("-*-*-*-*-*-*-*-*以下打勾選擇-*-*-*-*-*-*-*-*");
		String custom_no = "";
		ArrayList<String> selectedCustom = this.getSelectedCustom();

		switch (this.assignedProg) {

		case FancyGUI.C4_EXCEL_NOVARTIS:
			new Excel_C4_NOVARTIS().getExcel(selectedCustom);
			break;
		case FancyGUI.AIDC_EXCEL:
			new Excel_AIDC(isExpImp,rdbtnAIR.isSelected()).getExcel(selectedCustom);
			break;
		} // end switch/case
		
	}

	private void run_one_drno_to_one_output() throws Exception {
		System.out.println("-*-*-*-*-*-*-*-*以下打勾選擇-*-*-*-*-*-*-*-*");
		for (String[] drno_bk_pair : getSelectedDrnoBKnoPair()) {
			String custom_no = drno_bk_pair[0];
			String booking_no = drno_bk_pair[1];
			System.out.printf("Row %s is checked \n", custom_no);
			switch (this.assignedProg) {
			case FancyGUI.XML_WPG:
			case FancyGUI.AE_XML_WPG:
			case FancyGUI.AI_XML_WPG:

				Xml_WPG wpg = new Xml_WPG();
				wpg.getXML(custom_no);

				break;
			case FancyGUI.EXCEL_BENQ:
			case FancyGUI.AE_EXCEL_BENQ:

				Excel_AE_BENQ eBenq = new Excel_AE_BENQ();
				eBenq.getExcel(custom_no);

				break;
			// 2017-11-14 Jason Add generate AUO import XML
			case FancyGUI.BR_XML_AUO_IMP:
			case FancyGUI.BR_XML_AUO_EXP:
				if (isExpImp.equals("EXP")) {
					Xml_AUO_Export auoI = new Xml_AUO_Export(configPrefix);
					auoI.getXML(custom_no);
				} else {
					Xml_AUO_Import auoI = new Xml_AUO_Import(configPrefix);
					auoI.getXML(custom_no);
				}

				break;
			case FancyGUI.C1_EXCEL_BENQ:
				new Excel_C1_BENQ().getExcel(custom_no);
				break;
			case FancyGUI.C2_EXCEL_BENQ:
			case FancyGUI.C4_EXCEL_BENQ3:
				boolean isMerge = true; // 將多筆報單寫進同一個EXCEL 裡
				if (isMerge) {
					new Excel_C2_BENQ().getExcel(this.getSelectedCustom());

					// 強制結束
					return;
				} else {
					new Excel_C2_BENQ().getExcel(custom_no);
					break;
				}
			case FancyGUI.C3_EXCEL_BENQ:
				new Excel_C3_BENQ().getExcel(custom_no, booking_no);
				break;
			case FancyGUI.C3_EXCEL_COMMONUSE:
				new Excel_C3_CommonUse().getExcel(custom_no, booking_no);
				break;
			case FancyGUI.C4_EXCEL_BENQ:
				new Excel_C4_BENQ().getExcel(custom_no);
				break;
			case FancyGUI.C4_EXCEL_BENQ2:
				new Excel_C4_BENQ2().getExcel(custom_no);
				break;
		
			} // end switch/case
		} // end for(custom_no)
	}

	private ResultSet getCustomInfo(String sWhere, String searchMode) throws Exception {
		Connection conn = OutputCommon.connSQL();

		String exp_sql = "select DCL_DOC_NO as '報單號碼', MAWB, HAWB, " + " CONVERT(VARCHAR(5), DCL_DATE, 101) as '報關日期', "
				+ " RL_DATE as '放行日期', RL_TIME as '放行時間', " + " DOC_HEAD_DOC_NO as '業務單號' " + " from DOC_HEAD "
				+ " where 1=1 " + sWhere;

		String imp_sql = "select DCL_DOC_NO as '報單號碼', MAWB, HAWB, " + " CONVERT(VARCHAR(5), DCL_DATE, 101) as '報關日期', "
				+ " '20'+SUBSTRING(RL_DATE, 1, 2)+'/'+SUBSTRING(RL_DATE, 3, 2)+'/'+SUBSTRING(RL_DATE, 5, 2) as '放行日期', "
				+ " SUBSTRING(RL_DATE, 7, 2)+':'+SUBSTRING(RL_DATE, 9, 2) as '放行時間', " + " DOC_HEAD_DOC_NO as '業務單號' "
				+ " from DOC_H_I " + " where 1=1 " + sWhere;
		String sql;
		if (searchMode.equals("EXP"))
			sql = exp_sql;
		else
			sql = imp_sql;

		// sql = "select * from DOC_H_I where REPLACE(REPLACE(DCL_DOC_NO,'/',''),'
		// ','')='CBG2066060E030'";

		PreparedStatement ps = conn.prepareStatement(sql);
		// ps.setString(1, custom_no);

		return ps.executeQuery();
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
			columnNames.add("勾選");
			for (int column = 1; column <= columnCount; column++) {
				columnNames.add(metaData.getColumnName(column));
			}

			data = new Vector<Vector<Object>>();
			while (rs.next()) {
				Vector<Object> vector = new Vector<Object>();
				vector.add(false);
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

	/**
	 * JDatePicker 前面的Textfield 顯示方式
	 * 
	 * @author user
	 *
	 */
	@SuppressWarnings("serial")
	public class LabelDateFormatter extends AbstractFormatter {
		private String datePatern = "yyyy/MM/dd";

		private SimpleDateFormat dateFormatter = new SimpleDateFormat(datePatern);

		@Override
		public Object stringToValue(String text) throws ParseException {
			return dateFormatter.parseObject(text);
		}

		@Override
		public String valueToString(Object value) throws ParseException {
			if (value != null) {
				Calendar cal = (Calendar) value;
				return dateFormatter.format(cal.getTime());
			}

			return "";
		}
	}
}
