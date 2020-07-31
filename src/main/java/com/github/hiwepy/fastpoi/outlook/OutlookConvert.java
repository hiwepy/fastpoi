package com.github.hiwepy.fastpoi.outlook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.InputStream;

import org.apache.poi.hsmf.datatypes.AttachmentChunks;
import org.apache.poi.hsmf.extractor.OutlookTextExtactor;

import com.github.hiwepy.fastpoi.utils.ExtensionUtils;
import com.github.hiwepy.fastpoi.utils.LogUtils;
import com.github.hiwepy.fastpoi.utils.OutlookUtils;

public class OutlookConvert{

	private static OutlookConvert instance = null;
	private static Object initLock = new Object();
	
	private OutlookConvert(){};
	public static OutlookConvert getInstance(){
		if (instance == null) {
			synchronized(initLock) {
				instance= new OutlookConvert();
			}
		}
		return  instance;
	}
	
	public void outlook2Text(String inFilePath,String outFilePath) throws Exception {
		if(ExtensionUtils.isOutlook(inFilePath)){
			InputStream inputStream = null;
			try {
				File pptFile = new File(inFilePath);
				inputStream = new FileInputStream(pptFile);
				OutlookTextExtactor extactor = new OutlookTextExtactor(inputStream);
				FileWriter writer = null;
				try {
					writer = new FileWriter(new File(outFilePath));
					writer.write(extactor.getText().trim());
					String attDirName = outFilePath + "-att";
					AttachmentChunks[] attachments = extactor.getMAPIMessage().getAttachmentFiles();
					if(attachments.length > 0) {
						File d = new File(attDirName);
						if(d.mkdir()) {
							for(AttachmentChunks attachment : attachments) {
								OutlookUtils.processAttachment(attachment, d);
							}
						} else {
							LogUtils.error(new Exception("Can't create directory "+attDirName));
						}
					}
				} finally {
					if(writer != null) {
						try {
							writer.close();
						} catch (Exception e) {
							LogUtils.error(e);
						}
					}
				}
			} finally {
				if (inputStream != null) {
					try {
						inputStream.close();
					} catch (Exception e) {
						LogUtils.error(e);
					}
				}
			}
		}
	}

}
