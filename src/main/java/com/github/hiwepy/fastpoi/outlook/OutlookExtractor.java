package com.github.hiwepy.fastpoi.outlook;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.hsmf.MAPIMessage;
import org.apache.poi.hsmf.datatypes.AttachmentChunks;
import org.apache.poi.hsmf.extractor.OutlookTextExtactor;

import com.github.hiwepy.fastpoi.outlook.entity.OutlookEntity;
import com.github.hiwepy.fastpoi.utils.ExtensionUtils;

public class OutlookExtractor{

	private static OutlookExtractor instance = null;
	private static Object initLock = new Object();
	
	private OutlookExtractor(){};
	public static OutlookExtractor getInstance(){
		if (instance == null) {
			synchronized(initLock) {
				instance= new OutlookExtractor();
			}
		}
		return  instance;
	}
	
	public OutlookEntity getOutlook(String inFilePath) throws Exception {
		OutlookEntity temp = null;
		if(ExtensionUtils.isOutlook(inFilePath)){
			OutlookTextExtactor extactor =  OutlookReader.getInstance().getTextExtactor(inFilePath);
			SummaryInformation info = extactor.getSummaryInformation();
			MAPIMessage message = extactor.getMAPIMessage();
			temp = new OutlookEntity();
			temp.setTitle(info.getTitle());
			temp.setAuthor(info.getAuthor());
			temp.setSubject(message.getSubject());
			temp.setKeywords(info.getKeywords());
			temp.setTextBody(message.getTextBody().trim());
			temp.setCreateDateTime(info.getCreateDateTime());
			temp.setDisplayFrom(message.getDisplayFrom());
			temp.setDisplayTo(message.getDisplayTo());
			temp.setDisplayCC(message.getDisplayCC());
			temp.setDisplayBCC(message.getDisplayBCC());
			
			temp.setHeaders(message.getHeaders());
			temp.setRecipientEmailAddress(message.getRecipientEmailAddress());
			temp.setRecipientEmailAddressList(message.getRecipientEmailAddressList());
			temp.setRecipientNames(message.getRecipientNames());
			temp.setRecipientNamesList(message.getRecipientNamesList());
			temp.setRevNumber(info.getRevNumber());
			
			AttachmentChunks[] attachments = message.getAttachmentFiles();
			if(attachments.length > 0) {
				Map<String, byte[]> attachmentMaps = new HashMap<String, byte[]>();
				for(AttachmentChunks attachment : attachments) {
					String fileName = attachment.attachFileName.toString();
					if(attachment.attachLongFileName != null) {
						fileName = attachment.attachLongFileName.toString();
					}
					attachmentMaps.put(fileName, attachment.attachData.getValue());
				}
			}
		}
		return temp;
	}



}
