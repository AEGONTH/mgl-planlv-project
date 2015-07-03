package com.adms.mglplanreport.util;

import net.lingala.zip4j.core.ZipFile;
import net.lingala.zip4j.exception.ZipException;

public class ZipUtil {

	public static void extractAll(String source, String destination, String pwd) throws ZipException {
		ZipFile zipFile = new ZipFile(source);
		if(zipFile.isEncrypted()) {
			zipFile.setPassword(pwd);
		}
		zipFile.extractAll(destination);
	}
}
