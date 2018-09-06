package GUI;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Hashtable;
import java.util.Properties;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JFormattedTextField.AbstractFormatter;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextField;
import javax.swing.WindowConstants;
import org.jdatepicker.impl.JDatePanelImpl;
import org.jdatepicker.impl.JDatePickerImpl;
import org.jdatepicker.impl.UtilDateModel;

import OutputMethod.Excel_Common_OP_Key_Report;

import javax.swing.JRadioButton;
import javax.swing.ButtonGroup;

public class ReportConditionGUI {

	public static final int Common_Excel_OP_Key_Report = 8001; // "Excel_Common_OP_Key_Report"
	
	private String isExpImp = ""; // AE or AI
	private String sWhere = ""; // AE or AI
	boolean isShowSelectAll = true;

	private JFrame frame;
	private JLabel label_3;

	private JDatePickerImpl dpDrDate;
	private JDatePickerImpl dpDrDate2;
	private JLabel label_4;
	private JLabel lblOP;
	private JTextField tfOP;
	private JCheckBox bRelease;
	
	private final ButtonGroup buttonGroup = new ButtonGroup();
	
	JRadioButton rdbtnC1 = new JRadioButton("海出報單");
	JRadioButton rdbtnC2 = new JRadioButton("海進報單");
	JRadioButton rdbtnC3 = new JRadioButton("空出報單");
	JRadioButton rdbtnC4 = new JRadioButton("空進報單");

	private int gui_type;

	private Hashtable<String, String> conditionPack = new Hashtable<String, String>();

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					ReportConditionGUI window = new ReportConditionGUI();
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
	public ReportConditionGUI() {
		initialize();
	}
	
	public ReportConditionGUI(int gui_type) {
		this.gui_type = gui_type;
		
		initialize();

		switch (gui_type) {
		case Common_Excel_OP_Key_Report:
			frame.setTitle(Excel_Common_OP_Key_Report.programmeTitle);
			break;
		
		default:
			JOptionPane.showMessageDialog(null, "sub-Program 選擇器回傳問題", "錯誤 GUI-002", JOptionPane.INFORMATION_MESSAGE);
			break;
		}

		this.frame.setVisible(true);
	}

	
	/**
	 * Initialize the contents of the frame.
	 */
	@SuppressWarnings("serial")
	private void initialize() {
		frame = new JFrame();
		frame.setTitle("title set by program, if you see this, then there's problem");
		frame.setBounds(100, 100, 648, 227);
		frame.setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);
		frame.getContentPane().setLayout(null);

		Dimension dim = Toolkit.getDefaultToolkit().getScreenSize();
		frame.setLocation(dim.width / 2 - frame.getSize().width / 2,
				dim.height / 2 - frame.getSize().height / 2);

		JButton runBtn = new JButton("\u8F49\u51FA");
		runBtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				runXML();
			}
		});
		runBtn.setBounds(507, 125, 99, 27);
		frame.getContentPane().add(runBtn);

		label_3 = new JLabel("\u5831\u95DC\u65E5\u671F");
		label_3.setFont(new Font("新細明體", Font.PLAIN, 18));
		label_3.setBounds(14, 20, 80, 19);
		frame.getContentPane().add(label_3);

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
		dpDrDate.setLocation(96, 16);

		frame.getContentPane().add(dpDrDate);

		UtilDateModel model2 = new UtilDateModel();
		JDatePanelImpl datePanel2 = new JDatePanelImpl(model2, p);
		dpDrDate2 = new JDatePickerImpl(datePanel2, new LabelDateFormatter());
		dpDrDate2.getJFormattedTextField().setFont(new Font("新細明體", Font.PLAIN, 18));
		dpDrDate2.getJFormattedTextField().setBackground(Color.WHITE);
		dpDrDate2.setSize(116, 27);
		dpDrDate2.setLocation(239, 16);

		frame.getContentPane().add(dpDrDate2);

		label_4 = new JLabel("~");
		label_4.setFont(new Font("新細明體", Font.PLAIN, 18));
		label_4.setBounds(215, 19, 20, 19);
		frame.getContentPane().add(label_4);

		lblOP = new JLabel("OP");
		lblOP.setFont(new Font("新細明體", Font.PLAIN, 18));
		lblOP.setBounds(14, 55, 80, 19);
		frame.getContentPane().add(lblOP);

		tfOP = new JTextField();
		tfOP.setFont(new Font("新細明體", Font.PLAIN, 18));
		tfOP.setColumns(10);
		tfOP.setBounds(96, 52, 116, 25);
		frame.getContentPane().add(tfOP);

		bRelease = new JCheckBox("\u5DF2\u653E\u884C");
		bRelease.setSelected(true);
		bRelease.setFont(new Font("新細明體", Font.PLAIN, 18));
		bRelease.setBounds(240, 87, 115, 27);
		frame.getContentPane().add(bRelease);

		JButton clearConBtn = new JButton("清除條件");
		clearConBtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				clearCondition();
			}
		});
		clearConBtn.setBounds(14, 125, 99, 27);
		frame.getContentPane().add(clearConBtn);
		
		rdbtnC1.setFont(new Font("新細明體", Font.PLAIN, 18));
		buttonGroup.add(rdbtnC1);
		rdbtnC1.setBounds(384, 16, 113, 27);
		frame.getContentPane().add(rdbtnC1);
		
		rdbtnC2.setFont(new Font("新細明體", Font.PLAIN, 18));
		buttonGroup.add(rdbtnC2);
		rdbtnC2.setBounds(384, 51, 113, 27);
		frame.getContentPane().add(rdbtnC2);
		
		rdbtnC3.setFont(new Font("新細明體", Font.PLAIN, 18));
		buttonGroup.add(rdbtnC3);
		rdbtnC3.setBounds(384, 87, 113, 27);
		frame.getContentPane().add(rdbtnC3);
		
		rdbtnC4.setFont(new Font("新細明體", Font.PLAIN, 18));
		buttonGroup.add(rdbtnC4);
		rdbtnC4.setBounds(384, 119, 113, 27);
		frame.getContentPane().add(rdbtnC4);
	}

	public void clearCondition() {
		conditionPack.clear();
		
		bRelease.setSelected(true);
		dpDrDate.getJFormattedTextField().setValue(null);
		dpDrDate2.getJFormattedTextField().setValue(null);

		this.buttonGroup.clearSelection();
		
		tfOP.setText("");
	}

	public void getWhere() {
		conditionPack.clear();
		
		String sWhere = "";

		String sDrDt = dpDrDate.getJFormattedTextField().getText().trim();
		String sDrDt2 = dpDrDate2.getJFormattedTextField().getText().trim();
		if (!sDrDt.isEmpty()) {
			if (!sDrDt2.isEmpty()) {
				sWhere += " and DCL_DATE between '" + sDrDt + "' and '" + sDrDt2 + "' ";
				conditionPack.put("DCL_DATE", sDrDt+" ~ " + sDrDt2);
			} else {
				sWhere += " and DCL_DATE = '" + sDrDt + "' ";
				conditionPack.put("DCL_DATE", sDrDt);
			}
		}
		else conditionPack.put("DCL_DATE", "");

		String sOP = tfOP.getText().trim().toUpperCase();
		if (!sOP.isEmpty()) {
			sWhere += " and UPPER(OP_CODE) like '%" + sOP + "%' ";
			conditionPack.put("OP_CODE", sOP);
		}
		else conditionPack.put("OP_CODE", "");

		if (bRelease.isSelected()) {
			sWhere += " and RL_DATE is not null ";
			conditionPack.put("isRelease", "已放行");
		}
		else conditionPack.put("isRelease", "全部");
		
		// rdbtnSEA
		if (rdbtnC1.isSelected()) {
			sWhere += " and AIR_SEA = '1' ";
			this.isExpImp = "EXP";
			conditionPack.put("JOB_TYPE", "C1");
		}
		if(rdbtnC2.isSelected()) {
			sWhere += " and AIR_SEA = '1' ";
			this.isExpImp = "IMP";
			conditionPack.put("JOB_TYPE", "C2");
		}
		if (rdbtnC3.isSelected()) {
			sWhere += " and AIR_SEA = '4' ";
			this.isExpImp = "EXP";
			conditionPack.put("JOB_TYPE", "C3");
		}
		if(rdbtnC4.isSelected()) {
			sWhere += " and AIR_SEA = '4' ";
			this.isExpImp = "IMP";
			conditionPack.put("JOB_TYPE", "C4");
		}
	
		
		this.sWhere = sWhere;
	}

	

	public void runXML() {
		try {
			checkCondition();
			getWhere();
			
			switch (gui_type) {
			case Common_Excel_OP_Key_Report:
				Excel_Common_OP_Key_Report op_rpt = new Excel_Common_OP_Key_Report(sWhere, isExpImp, conditionPack);
				op_rpt.run();
				break;
			
			default:
				throw new Exception("抓不到報表代號");
			}
			
		} // end try/catch

		catch (Exception e) {
			e.printStackTrace();
			JOptionPane.showMessageDialog(null, "錯誤訊息 - " + e.getMessage(), "錯誤 GUI-RPT_CONDITION_GUI", JOptionPane.ERROR_MESSAGE);
		}
	}
	


	private void checkCondition() throws Exception {
		// TODO Auto-generated method stub
		if(rdbtnC1.isSelected() || rdbtnC2.isSelected() || rdbtnC3.isSelected() || rdbtnC4.isSelected())
		{
			// 有一個有選到
		}
		else {
			throw new Exception("業別未選擇");
		}
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
