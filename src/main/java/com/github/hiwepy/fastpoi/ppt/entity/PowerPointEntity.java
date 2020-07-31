package com.github.hiwepy.fastpoi.ppt.entity;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class PowerPointEntity {
	
	// 标题
	private String title;
	// 作者
	private String author;
	// 关键字
	private String keywords;
	// 创建时间
	private Date createDateTime;
	// 版本
	private String revNumber;
	
	
	private String headerText;
	private String footerText;
	private String dateTimeText;
	private int dateTimeFormat;
	
	private List<PowerPointSlideEntity> ppts = new ArrayList<PowerPointSlideEntity>(0);
	public PowerPointEntity(){}
	public PowerPointEntity(List<PowerPointSlideEntity> ppts){
		this.ppts = ppts;
	}
	
	public int size(){
		return this.ppts.size();
	}
	public String getHeaderText() {
		return headerText;
	}
	public void setHeaderText(String headerText) {
		this.headerText = headerText;
	}
	public String getFooterText() {
		return footerText;
	}
	public void setFooterText(String footerText) {
		this.footerText = footerText;
	}
	public String getDateTimeText() {
		return dateTimeText;
	}
	public void setDateTimeText(String dateTimeText) {
		this.dateTimeText = dateTimeText;
	}
	public int getDateTimeFormat() {
		return dateTimeFormat;
	}
	public void setDateTimeFormat(int dateTimeFormat) {
		this.dateTimeFormat = dateTimeFormat;
	}
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
