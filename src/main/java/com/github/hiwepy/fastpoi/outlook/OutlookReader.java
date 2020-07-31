package com.github.hiwepy.fastpoi.outlook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hsmf.extractor.OutlookTextExtactor;

import com.github.hiwepy.fastpoi.utils.ExtensionUtils;
import com.github.hiwepy.fastpoi.utils.LogUtils;

public class OutlookReader{


	private static OutlookReader instance = null;
	private static Object initLock = new Object();
	
	private OutlookReader(){};
	public static OutlookReader getInstance(){
		if (instance == null) {
			synchronized(initLock) {
				instance= new OutlookReader();
			}
		}
		return  instance;
	}
	
	public OutlookTextExtactor getTextExtactor(File file) throws IOException{
		if(file.exists()){
			if(ExtensionUtils.isOutlook(file.getAbsolutePath())){
				OutlookTextExtactor extactor  = null;
				InputStream inputStream = null;
				try {
					extactor = new OutlookTextExtactor(new FileInputStream(file));
				} finally {
					if (inputStream != null) {
						try {
							inputStream.close();
						} catch (Exception e) {
							LogUtils.error(e);
						}
					}
				}
				return  extactor;
			}else{
				throw new IOException("the file input is not a Outlook file ! ");
			}
		}else{
			throw new IOException("the file not exists ! ");
		}
	}
	
	public OutlookTextExtactor getTextExtactor(String inFilePath) throws IOException{
		return this.getTextExtactor(new File(inFilePath) );
	}

}
