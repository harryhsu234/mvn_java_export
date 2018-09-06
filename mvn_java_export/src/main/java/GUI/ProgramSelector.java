package GUI;

import java.awt.Dimension;
import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JLabel;
import java.awt.Font;
import java.awt.Toolkit;

import javax.swing.JButton;
import java.awt.event.ActionListener;
import java.io.IOException;
import java.awt.event.ActionEvent;
import javax.swing.JSeparator;
import javax.swing.SwingConstants;


import OutputMethod.*;
import net.lingala.zip4j.exception.ZipException;

public class ProgramSelector {

	public final static String programmeTitle = "轉檔程式選擇器";
	private JFrame frmTectpe;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					ProgramSelector window = new ProgramSelector();
					window.frmTectpe.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public ProgramSelector() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frmTectpe = new JFrame();
		frmTectpe.setResizable(false);
		frmTectpe.setTitle(this.programmeTitle);
		frmTectpe.setBounds(100, 100, 974, 561);
		frmTectpe.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frmTectpe.getContentPane().setLayout(null);
		
		Dimension dim = Toolkit.getDefaultToolkit().getScreenSize();
		frmTectpe.setLocation(dim.width/2-frmTectpe.getSize().width/2, dim.height/2-frmTectpe.getSize().height/2);
		
		JLabel label = new JLabel("\u7A7A\u904B\u51FA\u53E3");
		label.setHorizontalAlignment(SwingConstants.CENTER);
		label.setFont(new Font("新細明體", Font.PLAIN, 24));
		label.setBounds(26, 13, 147, 44);
		frmTectpe.getContentPane().add(label);
		
		JLabel label_1 = new JLabel("\u7A7A\u904B\u9032\u53E3");
		label_1.setHorizontalAlignment(SwingConstants.CENTER);
		label_1.setFont(new Font("新細明體", Font.PLAIN, 24));
		label_1.setBounds(417, 13, 147, 44);
		frmTectpe.getContentPane().add(label_1);
		
		JButton btnNewButton = new JButton("大聯大XML");
		btnNewButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				new FancyGUI(FancyGUI.AE_XML_WPG);
			}
		});
		btnNewButton.setBounds(26, 70, 147, 44);
		frmTectpe.getContentPane().add(btnNewButton);
		
		JButton btnNewButton_1 = new JButton("明碁 Excel (B2)");
		btnNewButton_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				new FancyGUI(FancyGUI.AE_EXCEL_BENQ);
			}
		});
		btnNewButton_1.setBounds(26, 127, 147, 44);
		frmTectpe.getContentPane().add(btnNewButton_1);
		
		JButton btnBenqGExcel = new JButton("BENQ G2 Excel");
		btnBenqGExcel.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				new FancyGUI(FancyGUI.C4_EXCEL_BENQ);
			}
		});
		btnBenqGExcel.setBounds(417, 127, 147, 44);
		frmTectpe.getContentPane().add(btnBenqGExcel);
		
		JButton btnBenqVExcel = new JButton("BENQ v2 Excel");
		btnBenqVExcel.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				new FancyGUI(FancyGUI.C4_EXCEL_BENQ2);
			}
		});
		btnBenqVExcel.setBounds(417, 184, 147, 44);
		frmTectpe.getContentPane().add(btnBenqVExcel);
		
		JButton btnNOVARTIS = new JButton("諾華 Excel");
		btnNOVARTIS.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				new FancyGUI(FancyGUI.C4_EXCEL_NOVARTIS);
			}
		});
		btnNOVARTIS.setBounds(417, 355, 147, 44);
		frmTectpe.getContentPane().add(btnNOVARTIS);
		
		JSeparator separator = new JSeparator();
		separator.setOrientation(SwingConstants.VERTICAL);
		separator.setBounds(384, 13, 19, 500);
		frmTectpe.getContentPane().add(separator);
		
		JButton button_1 = new JButton("\u5927\u806F\u5927XML");
		button_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				new FancyGUI(FancyGUI.AI_XML_WPG);
			}
		});
		button_1.setBounds(417, 70, 147, 44);
		frmTectpe.getContentPane().add(button_1);
		
		JButton btnAEMarvellWord = new JButton("Marvell Word");
		btnAEMarvellWord.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				new Excel_AE_Marvell().run();
			}
		});
		btnAEMarvellWord.setBounds(187, 70, 147, 44);
		frmTectpe.getContentPane().add(btnAEMarvellWord);
		
		JButton btnAEAdvancedRTF = new JButton("日月光 RTF");
		btnAEAdvancedRTF.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				new Excel_AE_AdvancedRTF().run();
			}
		});
		btnAEAdvancedRTF.setBounds(187, 127, 147, 44);
		frmTectpe.getContentPane().add(btnAEAdvancedRTF);
		
		JButton btnAUOXML_EXP = new JButton("AUO XML");
		btnAUOXML_EXP.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				new FancyGUI(FancyGUI.BR_XML_AUO_IMP);
			}
		});
		btnAUOXML_EXP.setBounds(417, 298, 147, 44);
		frmTectpe.getContentPane().add(btnAUOXML_EXP);
		
		JLabel label_2 = new JLabel("空運出口文件轉檔");
		label_2.setHorizontalAlignment(SwingConstants.CENTER);
		label_2.setFont(new Font("新細明體", Font.PLAIN, 14));
		label_2.setBounds(187, 39, 147, 18);
		frmTectpe.getContentPane().add(label_2);
		
		JSeparator separator_1 = new JSeparator();
		separator_1.setOrientation(SwingConstants.VERTICAL);
		separator_1.setBounds(578, 13, 19, 500);
		frmTectpe.getContentPane().add(separator_1);
		
		JButton btnC1BenqExcel = new JButton("BENQ Excel");
		btnC1BenqExcel.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				new FancyGUI(FancyGUI.C1_EXCEL_BENQ);
			}
		});
		btnC1BenqExcel.setBounds(611, 70, 147, 44);
		frmTectpe.getContentPane().add(btnC1BenqExcel);
		
		JLabel label_3 = new JLabel("海運出口");
		label_3.setHorizontalAlignment(SwingConstants.CENTER);
		label_3.setFont(new Font("新細明體", Font.PLAIN, 24));
		label_3.setBounds(611, 13, 147, 44);
		frmTectpe.getContentPane().add(label_3);
		
		JButton button_2 = new JButton("AUO XML");
		button_2.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				new FancyGUI(FancyGUI.BR_XML_AUO_EXP);
			}
		});
		button_2.setBounds(26, 184, 147, 44);
		frmTectpe.getContentPane().add(button_2);
		
		JSeparator separator_2 = new JSeparator();
		separator_2.setOrientation(SwingConstants.VERTICAL);
		separator_2.setBounds(772, 13, 19, 500);
		frmTectpe.getContentPane().add(separator_2);
		
		JLabel label_4 = new JLabel("海運進口");
		label_4.setHorizontalAlignment(SwingConstants.CENTER);
		label_4.setFont(new Font("新細明體", Font.PLAIN, 24));
		label_4.setBounds(611, 236, 147, 44);
		frmTectpe.getContentPane().add(label_4);
		
		JButton button_3 = new JButton("BENQ Excel");
		button_3.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				new FancyGUI(FancyGUI.C2_EXCEL_BENQ);
			}
		});
		button_3.setBounds(611, 295, 147, 44);
		frmTectpe.getContentPane().add(button_3);
		
		JButton btnDanliExcel = new JButton("丹利 Excel");
		btnDanliExcel.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				new Excel_AE_DANLI().run();
			}
		});
		btnDanliExcel.setBounds(187, 184, 147, 44);
		frmTectpe.getContentPane().add(btnDanliExcel);
		
		JButton btnAE_APEX_BJN_Excel = new JButton("精銳(重慶/北京)");
		btnAE_APEX_BJN_Excel.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				new Excel_AE_APEX_BJN().run();
			}
		});
		// btnAE_APEX_Excel.setEnabled(false);
		btnAE_APEX_BJN_Excel.setBounds(187, 238, 147, 44);
		frmTectpe.getContentPane().add(btnAE_APEX_BJN_Excel);
		
		JLabel label_5 = new JLabel("共用報表");
		label_5.setHorizontalAlignment(SwingConstants.CENTER);
		label_5.setFont(new Font("新細明體", Font.PLAIN, 24));
		label_5.setBounds(805, 13, 147, 44);
		frmTectpe.getContentPane().add(label_5);
		
		JButton btnOP_KEY_DR = new JButton("OP 打單排行");
//		btnOP_KEY_DR.setEnabled(false);
		btnOP_KEY_DR.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				new ReportConditionGUI(ReportConditionGUI.Common_Excel_OP_Key_Report);
			}
		});
		btnOP_KEY_DR.setBounds(805, 70, 147, 44);
		frmTectpe.getContentPane().add(btnOP_KEY_DR);
		
		JButton btnAE_APEX_SHA_Excel = new JButton("精銳(上海/廈門)");
		btnAE_APEX_SHA_Excel.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				new Excel_AE_APEX_SHA().run();
			}
		});
		btnAE_APEX_SHA_Excel.setBounds(187, 295, 147, 44);
		frmTectpe.getContentPane().add(btnAE_APEX_SHA_Excel);
		
		JButton btnC3_BENQ = new JButton("BENQ Excel");
		btnC3_BENQ.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				new FancyGUI(FancyGUI.C3_EXCEL_BENQ);
			}
		});
		btnC3_BENQ.setBounds(26, 238, 147, 44);
		frmTectpe.getContentPane().add(btnC3_BENQ);
		
		JButton button = new JButton("BENQ (海進版)");
		button.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				new FancyGUI(FancyGUI.C4_EXCEL_BENQ3);
			}
		});
		button.setBounds(417, 241, 147, 44);
		frmTectpe.getContentPane().add(button);
		
		JButton btnAE_KST = new JButton("世同(二併一)");
		btnAE_KST.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				new Excel_AE_KST().run();
			}
		});
		btnAE_KST.setBounds(187, 355, 147, 44);
		frmTectpe.getContentPane().add(btnAE_KST);
		
		JButton btnAIDC_Excel = new JButton("AIDC Excel");
		btnAIDC_Excel.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				new FancyGUI(FancyGUI.AIDC_EXCEL);
			}
		});
		btnAIDC_Excel.setBounds(805, 127, 147, 44);
		frmTectpe.getContentPane().add(btnAIDC_Excel);
		
		JButton button_4 = new JButton("通用版");
		button_4.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				new FancyGUI(FancyGUI.C3_EXCEL_COMMONUSE);
			}
		});
		button_4.setBounds(26, 295, 147, 44);
		frmTectpe.getContentPane().add(button_4);
		
		JButton btnAE_WPI_GROUP = new JButton("世平興業 ZIP");
		btnAE_WPI_GROUP.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					new Excel_AE_WPI_GROUP().run();
				} catch (IOException | ZipException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
			}
		});
		btnAE_WPI_GROUP.setBounds(187, 412, 147, 44);
		frmTectpe.getContentPane().add(btnAE_WPI_GROUP);
		
		JButton btnAE_BSI = new JButton("晟宇 Excel");
		btnAE_BSI.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				new Excel_AE_BSI().run();
			}
		});
		btnAE_BSI.setBounds(187, 469, 147, 44);
		frmTectpe.getContentPane().add(btnAE_BSI);
		
		JButton btnAE_LEA = new JButton("利益得文件整理");
		btnAE_LEA.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				new Excel_AE_LEA().run();
			}
		});
		btnAE_LEA.setBounds(26, 355, 147, 44);
		frmTectpe.getContentPane().add(btnAE_LEA);
		
		
		
		
	}
}
