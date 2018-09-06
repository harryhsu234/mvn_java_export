package GUI;

import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JTextField;
import javax.swing.WindowConstants;
import Util.Config;

import javax.swing.JLabel;
import java.util.Hashtable;
import javax.swing.JButton;
import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;

public class ParameterSetting {

	private JFrame iniFrame;
	private JTextField fromContact;
	private JTextField fromEmail;
	private JTextField fromTel;
	private JTextField fromDuns;
	private JTextField toContact;
	private JTextField toEmail;
	private JTextField toTel;
	private JTextField toDuns;
	private String type;
	private JTextField toBroker;
	
	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					ParameterSetting window = new ParameterSetting();
					window.iniFrame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}
	
	/**
	 * Create the application.
	 */
	public ParameterSetting() {
		initialize();
		
	}
	
	/**
	 * Create the application.
	 */
	public ParameterSetting(String type) throws Exception {
		initialize();
		this.type = type;
		
		Hashtable<String,String> formRole = new Config().getConfig(type+"_FROM");
		fromContact.setText(formRole.get("contactName"));
		fromEmail.setText(formRole.get("EmailAddress"));
		fromTel.setText(formRole.get("CommunicationsNumber"));
		fromDuns.setText(formRole.get("GlobalBusinessIdentifier"));
		
		Hashtable<String,String> toRole = new Config().getConfig(type+"_TO");
		toContact.setText(toRole.get("contactName"));
		toEmail.setText(toRole.get("EmailAddress"));
		toTel.setText(toRole.get("CommunicationsNumber"));
		toDuns.setText(toRole.get("GlobalBusinessIdentifier"));
		toBroker.setText(toRole.get("Broker"));
		
		this.iniFrame.setVisible(true);
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		iniFrame = new JFrame();
		iniFrame.getContentPane().setFont(iniFrame.getContentPane().getFont().deriveFont(12f));
		iniFrame.setTitle("參數設定");
		iniFrame.setBounds(100, 100, 451, 472);
		iniFrame.setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);
		iniFrame.getContentPane().setLayout(null);
		
		fromContact = new JTextField();
		fromContact.setBounds(216, 59, 197, 21);
		iniFrame.getContentPane().add(fromContact);
		fromContact.setColumns(10);
		
		JLabel lblFormRole = new JLabel("From Role");
		lblFormRole.setBounds(59, 24, 59, 15);
		iniFrame.getContentPane().add(lblFormRole);
		
		JLabel lblNewLabel = new JLabel("contactName");
		lblNewLabel.setBounds(126, 56, 65, 27);
		iniFrame.getContentPane().add(lblNewLabel);
		
		JLabel lblEmail = new JLabel("Email");
		lblEmail.setBounds(126, 90, 65, 27);
		iniFrame.getContentPane().add(lblEmail);
		
		fromEmail = new JTextField();
		fromEmail.setColumns(10);
		fromEmail.setBounds(216, 93, 197, 21);
		iniFrame.getContentPane().add(fromEmail);
		
		fromTel = new JTextField();
		fromTel.setColumns(10);
		fromTel.setBounds(216, 130, 197, 21);
		iniFrame.getContentPane().add(fromTel);
		
		JLabel lblTel = new JLabel("TEL");
		lblTel.setBounds(126, 127, 65, 27);
		iniFrame.getContentPane().add(lblTel);
		
		JLabel lblDnus = new JLabel("DUNS");
		lblDnus.setBounds(126, 161, 65, 27);
		iniFrame.getContentPane().add(lblDnus);
		
		fromDuns = new JTextField();
		fromDuns.setColumns(10);
		fromDuns.setBounds(216, 164, 197, 21);
		iniFrame.getContentPane().add(fromDuns);
		
		JLabel label = new JLabel("contactName");
		label.setBounds(126, 241, 65, 27);
		iniFrame.getContentPane().add(label);
		
		toContact = new JTextField();
		toContact.setColumns(10);
		toContact.setBounds(216, 244, 197, 21);
		iniFrame.getContentPane().add(toContact);
		
		toEmail = new JTextField();
		toEmail.setColumns(10);
		toEmail.setBounds(216, 278, 197, 21);
		iniFrame.getContentPane().add(toEmail);
		
		JLabel label_1 = new JLabel("Email");
		label_1.setBounds(126, 275, 65, 27);
		iniFrame.getContentPane().add(label_1);
		
		JLabel lblToRole = new JLabel("To Role");
		lblToRole.setBounds(59, 209, 59, 15);
		iniFrame.getContentPane().add(lblToRole);
		
		JLabel label_3 = new JLabel("TEL");
		label_3.setBounds(126, 312, 65, 27);
		iniFrame.getContentPane().add(label_3);
		
		toTel = new JTextField();
		toTel.setColumns(10);
		toTel.setBounds(216, 315, 197, 21);
		iniFrame.getContentPane().add(toTel);
		
		toDuns = new JTextField();
		toDuns.setColumns(10);
		toDuns.setBounds(216, 349, 197, 21);
		iniFrame.getContentPane().add(toDuns);
		
		JLabel label_4 = new JLabel("DUNS");
		label_4.setBounds(126, 346, 65, 27);
		iniFrame.getContentPane().add(label_4);
		
		JButton btnNewButton = new JButton("Save");
		btnNewButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				save();
			}
		});
		btnNewButton.setBounds(326, 20, 87, 23);
		iniFrame.getContentPane().add(btnNewButton);
		
		JLabel lblBroker = new JLabel("Broker");
		lblBroker.setBounds(126, 383, 65, 27);
		iniFrame.getContentPane().add(lblBroker);
		
		toBroker = new JTextField();
		toBroker.setColumns(10);
		toBroker.setBounds(216, 386, 96, 21);
		iniFrame.getContentPane().add(toBroker);
	
	}
	
	/**
	 * save form to ini file.
	 * @throws Exception 
	 */
	private void save() {
		Config conf = new Config();
		Hashtable ht = new Hashtable();
		
		try {
			ht.put("contactName", fromContact.getText());
			ht.put("EmailAddress", fromEmail.getText());
			ht.put("CommunicationsNumber", fromTel.getText());
			ht.put("GlobalBusinessIdentifier", fromDuns.getText());
			conf.writeSection(type+"_FROM", ht);
			
			ht.put("contactName", toContact.getText());
			ht.put("EmailAddress", toEmail.getText());
			ht.put("CommunicationsNumber", toTel.getText());
			ht.put("GlobalBusinessIdentifier", toDuns.getText());
			ht.put("Broker", toBroker.getText());
			conf.writeSection(type+"_TO", ht);
			
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		
	}
}
