package com.adms.mglplanreport.app;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.io.InputStream;
import java.text.ParseException;
import java.util.Calendar;
import java.util.Vector;

import com.adms.mglplanreport.util.ZipUtil;
import com.adms.support.FileWalker;
import com.adms.utils.DateUtil;
import com.adms.utils.FileUtil;
import com.pff.PSTAttachment;
import com.pff.PSTException;
import com.pff.PSTFile;
import com.pff.PSTFolder;
import com.pff.PSTMessage;

public class GetZipFromEmail {

	private static int depth = -1;

	public static void main(String[] args) {

		try {
			System.out.println("====== Start ======");
			
//			<!-- Getting zip files from email archive(PST file) -->
//			System.out.println("====== Email ======");
			String outDir = "D:/project/reports/MGL/zip/201505";
			PSTFile pstFile = new PSTFile("D:/Email/Archive_PataweeCha_2015.pst");
			System.out.println(pstFile.getMessageStore().getDisplayName());
//			processPST(pstFile.getRootFolder(), outDir);

//			<!-- UnZip -->
			System.out.println("====== Unzip ======");
//			String zipDir = new String(outDir);
			String zipDir = "D:/project/upload file/auto report/zip";
			String destination = "D:/project/upload file/auto report/201505";
			FileWalker fw = new FileWalker();
			fw.walk(zipDir, new FilenameFilter() {
				
				@Override
				public boolean accept(File dir, String name) {
					return !name.contains("archive")
							&& name.endsWith(".zip");
				}
			});
			
			for(String file : fw.getFileList()) {
				System.out.println("File: " + file);
				try{
					if(file.contains("OTO")) {
						ZipUtil.extractAll(file, destination + "/OTO", passwordResolver(file));
					} else if(file.contains("TELE")) {
						String specific = destination + "/TELE";
						if(file.contains("MSIG_UOB")) {
							specific += "/AUTO_MSIGUOB";
						} else if(file.contains("MTI")) {
							specific += "/AUTO_MTIKBANK";
						} else if(file.contains("MTL")) {
							specific += "/AUTO_MTL";
						}
						
						ZipUtil.extractAll(file, specific, passwordResolver(file));
						File move = new File(file);
						FileUtil.getInstance().moveFile(file, move.getParent() + "/archive/" + move.getName());
					}
				} catch(Exception e) {
					e.printStackTrace();
				}
			}
			
			System.out.println("====== Finish ======");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	private static String passwordResolver(String name) throws ParseException {
		String OTOPWD = "aeg@n#yyyy", TELEPWD = "Aeg@nReport#yyyyMMdd";
		
		if(name.contains("OTO")) {
			return OTOPWD.replaceAll("#yyyy", name.substring(name.indexOf("_Report") - 4, name.indexOf("_Report")));
		} else if(name.contains("TELE")) {
			Calendar calendar = DateUtil.getCurrentCalendar();
			if(name.contains("ADAMS_")) {
				calendar.setTime(DateUtil.convStringToDate("yyMMdd", name.substring(name.indexOf(".zip") - 6, name.indexOf(".zip"))));
			} else {
				calendar.setTime(DateUtil.convStringToDate("ddMMyyyy", name.substring(name.indexOf(".zip") - 8, name.indexOf(".zip"))));
			}
			
			DateUtil.addDay(calendar, - 1);
			return TELEPWD.replaceAll("#yyyyMMdd", DateUtil.convDateToString("yyyyMMdd", calendar.getTime()).replaceAll("0", "@"));
		}
		
		return "";
	}

	private static void processPST(PSTFolder folder, String outDir) throws PSTException, IOException {
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
				processPST(childFolder, outDir);
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
//						@tele-intel.com for TELE, @onetoonecontacts.com for OTO
						System.out.println("Reading Email: " + email.getSubject() + " | from: " + email.getSenderEmailAddress());
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
