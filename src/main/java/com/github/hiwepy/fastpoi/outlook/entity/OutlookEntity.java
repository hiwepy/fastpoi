package com.github.hiwepy.fastpoi.outlook.entity;

import java.util.Date;
import java.util.Map;

public class OutlookEntity {

	// 标题
	private String title;
	// 作者
	private String author;
	// 主题
	private String subject;
	// 关键字
	private String keywords;
	// 内容
	private String textBody;
	// 附件内容
	private Map<String, byte[]> attachments;
	// 创建时间
	private Date createDateTime;
	// 发件人
	private String displayFrom;
	// 收件人
	private String displayTo;
	// 抄送人
	private String displayCC;
	// 暗送人
	private String displayBCC;

	private String[] headers;
	private String recipientEmailAddress;
	private String[] recipientEmailAddressList;
	private String recipientNames;
	private String[] recipientNamesList;

	
	private String revNumber;

	public String getRevNumber() {
		return revNumber;
	}

	public void setRevNumber(String revNumber) {
		this.revNumber = revNumber;
	}

	public String getTitle() {
		return title;
	}

	public void setTitle(String title) {
		this.title = title;
	}

	public String getSubject() {
		return subject;
	}

	public void setSubject(String subject) {
		this.subject = subject;
	}

	public String getAuthor() {
		return author;
	}

	public void setAuthor(String author) {
		this.author = author;
	}

	public String getKeywords() {
		return keywords;
	}

	public void setKeywords(String keywords) {
		this.keywords = keywords;
	}

	public String getTextBody() {
		return textBody;
	}

	public void setTextBody(String textBody) {
		this.textBody = textBody;
	}

	public Map<String, byte[]> getAttachments() {
		return attachments;
	}

	public void setAttachments(Map<String, byte[]> attachments) {
		this.attachments = attachments;
	}

	public Date getCreateDateTime() {
		return createDateTime;
	}

	public void setCreateDateTime(Date createDateTime) {
		this.createDateTime = createDateTime;
	}

	public String getDisplayFrom() {
		return displayFrom;
	}

	public void setDisplayFrom(String displayFrom) {
		this.displayFrom = displayFrom;
	}

	public String getDisplayTo() {
		return displayTo;
	}

	public void setDisplayTo(String displayTo) {
		this.displayTo = displayTo;
	}

	public String getDisplayCC() {
		return displayCC;
	}

	public void setDisplayCC(String displayCC) {
		this.displayCC = displayCC;
	}

	public String getDisplayBCC() {
		return displayBCC;
	}

	public void setDisplayBCC(String displayBCC) {
		this.displayBCC = displayBCC;
	}

	public String getRecipientEmailAddress() {
		return recipientEmailAddress;
	}

	public void setRecipientEmailAddress(String recipientEmailAddress) {
		this.recipientEmailAddress = recipientEmailAddress;
	}

	public String[] getRecipientEmailAddressList() {
		return recipientEmailAddressList;
	}

	public void setRecipientEmailAddressList(String[] recipientEmailAddressList) {
		this.recipientEmailAddressList = recipientEmailAddressList;
	}

	public String getRecipientNames() {
		return recipientNames;
	}

	public void setRecipientNames(String recipientNames) {
		this.recipientNames = recipientNames;
	}

	public String[] getRecipientNamesList() {
		return recipientNamesList;
	}

	public void setRecipientNamesList(String[] recipientNamesList) {
		this.recipientNamesList = recipientNamesList;
	}

	public String[] getHeaders() {
		return headers;
	}

	public void setHeaders(String[] headers) {
		this.headers = headers;
	}

}
