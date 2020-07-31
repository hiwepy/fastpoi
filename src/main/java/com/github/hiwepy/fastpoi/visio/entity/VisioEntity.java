package com.github.hiwepy.fastpoi.visio.entity;

import java.util.Date;

public class VisioEntity {
	// 标题
	private String title;
	// 作者
	private String author;
	// 关键字
	private String keywords;
	// 内容
	private String text;
	// 内容
	private String[] allTtext;
	// 创建时间
	private Date createDateTime;
	// 版本
	private String revNumber;

	public String getTitle() {
		return title;
	}

	public void setTitle(String title) {
		this.title = title;
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

	public String getText() {
		return text;
	}

	public void setText(String text) {
		this.text = text;
	}

	public String[] getAllTtext() {
		return allTtext;
	}

	public void setAllTtext(String[] allTtext) {
		this.allTtext = allTtext;
	}

	public Date getCreateDateTime() {
		return createDateTime;
	}

	public void setCreateDateTime(Date createDateTime) {
		this.createDateTime = createDateTime;
	}

	public String getRevNumber() {
		return revNumber;
	}

	public void setRevNumber(String revNumber) {
		this.revNumber = revNumber;
	}
		
}
