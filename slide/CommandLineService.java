package com.ebdesk.report.slide;
import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.lang.reflect.Field;

public class CommandLineService {	
	public CommandLineService() {
		
	}
	
	public int execute(String cmd) {
		try {
			Runtime rt = Runtime.getRuntime();
			Process pr = rt.exec(cmd);
			BufferedReader input = new BufferedReader(new InputStreamReader(pr.getInputStream()));
			String line = null;
			while((line=input.readLine()) != null) {
                System.out.println("Runnable :"+line);
            }
			Field field = pr.getClass().getDeclaredField("pid");
			field.setAccessible(true);
            int exitVal = pr.waitFor();
            System.out.println("Exited with error code "+exitVal);            
            return exitVal;
		} catch (Exception e) {
			e.printStackTrace();			
			return -1;
		}
	}
}