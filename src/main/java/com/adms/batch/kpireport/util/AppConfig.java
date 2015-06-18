package com.adms.batch.kpireport.util;

import org.springframework.context.ApplicationContext;
import org.springframework.context.support.ClassPathXmlApplicationContext;

public class AppConfig {
	
	private static AppConfig instance;
	
	private ApplicationContext context;
	
	public AppConfig() {
		context = new ClassPathXmlApplicationContext("/config/application-context-kpi-report.xml");
	}
	
	public static AppConfig getInstance() {
		if(instance == null) {
			instance = new AppConfig();
		}
		return instance;
	}
	
	public Object getBean(String bean) {
		return context.getBean(bean);
	}

}
