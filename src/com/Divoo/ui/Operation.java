package com.Divoo.ui;

import java.awt.BorderLayout;
import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;
import javax.swing.JLabel;
import javax.swing.JButton;

import com.Divoo.Extractor.Extractor;

import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;
import java.io.FileNotFoundException;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;

public class Operation extends JFrame
{

	private JPanel contentPane;
	
	/**
	 * Create the frame.
	 */
	public Operation()
	{
		setTitle("Mail Extractor");
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 450, 300);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);
		
		JLabel lblByMohamedMayla = new JLabel("By: Mohamed Mayla");
		lblByMohamedMayla.setBounds(259, 237, 177, 15);
		contentPane.add(lblByMohamedMayla);
		
		JButton btnStart = new JButton("Start");
		btnStart.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try
				{
				String path = System.getProperty("user.dir");
				Extractor ext = new Extractor();
				ext.readAllWordFileAt(path);
				}
				catch(Exception ex)
				{
					try
					{
						PrintWriter writer = new PrintWriter("Extracted Mails.txt", "UTF-8");
						writer.println(ex.getStackTrace().toString());
					} catch (FileNotFoundException e1)
					{
						e1.printStackTrace();
					} catch (UnsupportedEncodingException e1)
					{
						// TODO Auto-generated catch block
						e1.printStackTrace();
					}
					JOptionPane.showMessageDialog(null, "Error: please send the error.log file's content to the developer");
				}
				
				JOptionPane.showMessageDialog(null, "Done");
			}
		});
		btnStart.setBounds(12, 125, 424, 72);
		contentPane.add(btnStart);
		
		JLabel lblPutTheProgram = new JLabel("Put the program in the word files's directory");
		lblPutTheProgram.setBounds(12, 0, 387, 15);
		contentPane.add(lblPutTheProgram);
		
		JLabel lblPressStartButton = new JLabel("Press start button");
		lblPressStartButton.setBounds(12, 27, 387, 15);
		contentPane.add(lblPressStartButton);
		
		JLabel lblextractedMailWill = new JLabel("'Extracted Mail' will appear");
		lblextractedMailWill.setBounds(12, 54, 387, 15);
		contentPane.add(lblextractedMailWill);
		
		JLabel lblTheProgramOnly = new JLabel("The program only read doc and docx files");
		lblTheProgramOnly.setBounds(12, 81, 387, 15);
		contentPane.add(lblTheProgramOnly);
	}
}
