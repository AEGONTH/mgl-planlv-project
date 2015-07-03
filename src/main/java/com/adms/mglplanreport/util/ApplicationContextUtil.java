package com.adms.mglplanreport.util;

import org.springframework.context.ApplicationContext;
import org.springframework.context.support.ClassPathXmlApplicationContext;

public class ApplicationContextUtil {

	private final static String contextPath = "/config/application-context-mgl-report.xml";
	
	private static ApplicationContext ctx;
	
	public static ApplicationContext getApplicationContext() {
		if(ctx == null) {
			ctx = new ClassPathXmlApplicationContext(contextPath);
		}
		return ctx;
	}
}
