package search.excel;

import java.awt.Button;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Font;
import java.awt.Frame;
import java.awt.Label;
import java.awt.TextField;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;

import javax.swing.JFrame;

public class SearchExcelFrame extends Frame implements ActionListener{
	 private Label lbl_path;    // Declare a Label component 
	   private TextField tf_path; // Declare a TextField component 
	   private Button btn_submit;   // Declare a Button component
	   private Label lbl_array;
	   private TextField tf_array;
	private static final long serialVersionUID = 1L;

	public SearchExcelFrame() {
		setLayout(new FlowLayout());
		
		lbl_path = new Label("File Folder path");
		lbl_path.setPreferredSize(new Dimension(150,200));
		lbl_path.setFont(new Font("TimesRoman",0,20));
		add(lbl_path);
		
		tf_path = new TextField(80);
		tf_path.setPreferredSize(new Dimension(5000, 500));
		add(tf_path);
		
		lbl_array = new Label("Input search content, split by comma ','");
		lbl_array.setFont(new Font("TimesRoman",0,20));
		add(lbl_array);
		
		tf_array = new TextField(100);
		tf_array.setSize(2000, 400);
		add(tf_array);
		
		btn_submit = new Button("Search and Generate file");
		btn_submit.addActionListener(this);
		
		add(btn_submit);
		setTitle("Excel Stalker");
		setSize(1000,400);
		setVisible(true);
		
		
	}
	
	
	
	@Override
	public void actionPerformed(ActionEvent arg0) {
		String[] arr = tf_array.getText().toString().split(",");
		new SearchExcel().printExcel(new SearchExcel().getFileList(tf_path.getText()), arr);
	}
	public static void main(String[] args) {
		new SearchExcelFrame().addWindowListener(new WindowAdapter() {

	        public void windowClosing(WindowEvent evt) {
	            System.exit(0);
	        }

	});
		
	}
}
