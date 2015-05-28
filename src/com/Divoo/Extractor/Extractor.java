package com.Divoo.Extractor;

import java.io.*;
import java.util.HashSet;
import java.util.Set;
import java.util.Vector;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;

public class Extractor
{
	private Set emails;
	private String[] fileData;

	public Extractor()
	{
		// TODO Auto-generated constructor stub
		emails = new HashSet();
	}

	private String cleanString(String str, int dir) // dir 0 = front else rear
	{
		if (dir == 0)
		{
			for (int i = 0; i < str.length(); i++)
			{
				if (!Character.isLetterOrDigit(str.charAt(i))
						&& str.charAt(i) != '.')
				{
					// System.out.println("error at " + str.charAt(i));
					return str.substring(0, i);
				}
			}
		} else
		{
			for (int i = str.length() - 1; i >= 0; i--)
			{
				if (!Character.isLetterOrDigit(str.charAt(i))
						&& str.charAt(i) != '.' && str.charAt(i) != '-'
						&& str.charAt(i) != '_')
				{
					// System.out.println("error at " + str.charAt(i));
					return str.substring(i + 1, str.length());
				}
			}
		}
		return str;
	}

	private String cleanEmail(String mail, int atIndx)
	{
		String beforeAt = mail.substring(0, atIndx);
		String afterAt = mail.substring(atIndx + 1, mail.length());

		beforeAt = cleanString(beforeAt, 1);
		afterAt = cleanString(afterAt, 0);

		if (beforeAt.length() > 0 && afterAt.length() > 0)
			return beforeAt + "@" + afterAt;
		else
			return "null";

	}

	private void extract()
	{
		String tmpMail;
		for (int i = 0; i < fileData.length; i++)
		{
			int indx = fileData[i].indexOf('@');
			if (indx >= 0)
			{
				tmpMail = cleanEmail(fileData[i], indx);
				if (!tmpMail.equals("null"))
					emails.add(tmpMail);
			}
		}
	}

	public void readFile(String path)
	{
		File file = null;
		WordExtractor extractor = null;
		try
		{
			file = new File(path);
			FileInputStream fis = new FileInputStream(file.getAbsolutePath());
			HWPFDocument document = new HWPFDocument(fis);
			extractor = new WordExtractor(document);
			fileData = extractor.getParagraphText();
		} catch (Exception exep)
		{
			exep.printStackTrace();
		}
	}

	public void readDocxFile(String path)
	{
		File file = null;
		XWPFWordExtractor extractor = null;
		try
		{
			file = new File(path);
			FileInputStream fis = new FileInputStream(file.getAbsolutePath());
			XWPFDocument document = new XWPFDocument(fis);
			extractor = new XWPFWordExtractor(document);
			fileData = extractor.getText().split(" ");
		} catch (Exception exep)
		{
			exep.printStackTrace();
		}
	}

	public void readAllWordFileAt(String path)
	{
		File folder = new File(path);
		File[] listOfFiles = folder.listFiles();

		for (int i = 0; i < listOfFiles.length; i++)
		{
			if (listOfFiles[i].isFile())
			{
				String ext = listOfFiles[i].getName();
				ext = ext.substring(ext.lastIndexOf('.') + 1, ext.length());

				if (ext.equals("doc"))
				{
					System.out.println("#### begin - "
							+ listOfFiles[i].getName());
					readFile(path + "/" + listOfFiles[i].getName());
					extract();
				} else if (ext.equals("docx"))
				{
					System.out.println("#### begin - "
							+ listOfFiles[i].getName());
					readDocxFile(path + "/" + listOfFiles[i].getName());
					extract();
				}
			}
		}
		
		writeResultsToFile();
	}

	
	private void writeResultsToFile()
	{
		try
		{
			PrintWriter writer = new PrintWriter("Extracted Mails.txt", "UTF-8");
			String separator = "; ";
			
			
			writer.println(emails.size());
			for (Object str : emails)
			 {
				//System.out.println(str+separator);
				writer.print(str+separator);
			 }
			writer.close();
			
		} catch (FileNotFoundException | UnsupportedEncodingException e)
		{
			e.printStackTrace();
		}
		
		
	}
	
	public static void _main(String[] args)
	{
		//System.out.println("Working Directory = " + System.getProperty("user.dir"));
		Extractor ext = new Extractor();
		ext.readAllWordFileAt("/home/divoo/Desktop/emailextractor");
		
		String str ="null";
		switch (str)
		{
		case "Hi":
			break;
		}

	}
}
