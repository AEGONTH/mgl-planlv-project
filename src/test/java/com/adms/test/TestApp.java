package com.adms.test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Vector;

import com.pff.PSTAttachment;
import com.pff.PSTException;
import com.pff.PSTFile;
import com.pff.PSTFolder;
import com.pff.PSTMessage;

public class TestApp {

	private static int depth = -1;

	public static void main(String[] args) {

		try {
			System.out.println("start");
			String outDir = "D:/project/reports/MGL/in/zip/201505";
			File outfile = new File(outDir);
			if(!outfile.exists()) outfile.mkdirs();

			PSTFile pstFile = new PSTFile("D:/Email/Archive_PataweeCha_2015.pst");

			System.out.println(pstFile.getMessageStore().getDisplayName());

			processFolder(pstFile.getRootFolder(), outDir);

			System.out.println("finish");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private static void processFolder(PSTFolder folder, String outDir) throws PSTException,
			IOException {
		depth++;
		// the root folder doesn't have a display name
		if (depth > 0) {
			printDepth();
			System.out.println(folder.getDisplayName());
		}

		// go through the folders...
		if (folder.hasSubfolders()) {
			Vector<PSTFolder> childFolders = folder.getSubFolders();
			for (PSTFolder childFolder : childFolders) {
				processFolder(childFolder, outDir);
			}
		}

		// and now the emails for this folder
		if(folder.getDisplayName().equalsIgnoreCase("daily report")
				|| folder.getDisplayName().equalsIgnoreCase("autorpt")) {
			if (folder.getContentCount() > 0) {
				depth++;
				PSTMessage email = (PSTMessage) folder.getNextChild();
				while (email != null) {
					if(email.getSubject().toLowerCase().contains("confirmrpt")
							|| email.getSubject().toLowerCase().contains(" : report")
							|| (email.getSubject().toLowerCase().contains("autorpt_")
									&& !email.getSubject().toLowerCase().contains("app")
									&& !email.getSubject().toLowerCase().contains("yesfiles"))) {
						
						printDepth();
						System.out.println("Email: " + email.getSubject() + " | from: " + email.getSenderEmailAddress());
						
//						@tele-intel.com for TELE, @onetoonecontacts.com for OTO
						getAttachments(email, outDir + "/" + (email.getSenderEmailAddress().endsWith("@tele-intel.com") ? "TELE" : "OTO"));
					}
					email = (PSTMessage) folder.getNextChild();
				}
				depth--;
			}
		}
		depth--;
	}
	
	private static void getAttachments(PSTMessage email, String outDir) {
		try {
			File f = new File(outDir);
			if(!f.exists()) f.mkdirs();
			int numberOfAttachments = email.getNumberOfAttachments();
			for (int x = 0; x < numberOfAttachments; x++) {
				PSTAttachment attach = email.getAttachment(x);
				InputStream attachmentStream = attach.getFileInputStream();
				// both long and short filenames can be used for attachments
				String filename = attach.getLongFilename();
				if (filename.isEmpty()) {
					filename = attach.getFilename();
				}
				
				if(filename.toLowerCase().contains(".zip")) {
					FileOutputStream out = new FileOutputStream(outDir + "/" + filename);
					// 8176 is the block size used internally and should give the
					// best performance
					int bufferSize = 8176;
					byte[] buffer = new byte[bufferSize];
					int count = attachmentStream.read(buffer);
					while (count == bufferSize) {
						out.write(buffer);
						count = attachmentStream.read(buffer);
					}
					byte[] endBuffer = new byte[count];
					System.arraycopy(buffer, 0, endBuffer, 0, count);
					out.write(endBuffer);
					out.close();
				}
				
				attachmentStream.close();
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void printDepth() {
		for (int x = 0; x < depth - 1; x++) {
			System.out.print(" | ");
		}
		System.out.print(" |- ");
	}

}
