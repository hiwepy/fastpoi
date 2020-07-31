/*
 * Copyright (c) 2018, hiwepy (https://github.com/hiwepy).
 *
 * Licensed under the Apache License, Version 2.0 (the "License"); you may not
 * use this file except in compliance with the License. You may obtain a copy of
 * the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS, WITHOUT
 * WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the
 * License for the specific language governing permissions and limitations under
 * the License.
 */
package com.github.hiwepy.fastpoi.ppt;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hslf.model.HeadersFooters;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import com.github.hiwepy.fastpoi.ppt.entity.PowerPointEntity;
import com.github.hiwepy.fastpoi.utils.ExtensionUtils;

public class PowerPointExtractor {


	public PowerPointEntity getPowerPoint(String inFilePath)
			throws Exception {
		PowerPointEntity temp = null;
		if(ExtensionUtils.isPPt(inFilePath)){
			InputStream inputStream = null;
			try {
				File pptFile = new File(inFilePath);
				inputStream = new FileInputStream(pptFile);
				POIFSFileSystem fs = new POIFSFileSystem(inputStream); 
				if(ExtensionUtils.isPPt(inFilePath)){
					temp = new PowerPointEntity();
					
					HSLFSlideShow ppt = new HSLFSlideShow(fs);
					
					HeadersFooters hf = ppt.getNotesHeadersFooters();
					temp.setHeaderText(hf.getHeaderText());
					temp.setFooterText(hf.getFooterText());
					temp.setDateTimeText(hf.getDateTimeText());
					temp.setDateTimeFormat(hf.getDateTimeFormat());
					/*
					SummaryInformation info = ppt.getSummaryInformation();
					temp.setTitle(info.getTitle());
					temp.setAuthor(info.getAuthor());
					temp.setKeywords(info.getKeywords());
					temp.setCreateDateTime(info.getCreateDateTime());
					temp.setRevNumber(info.getRevNumber());*/
					
				}
			} finally {
				if (inputStream != null) {
					try {
						inputStream.close();
					} catch (Exception e) {
						System.err.println(e);
					}
				}
			}
		}
		return temp;
	}
	public PowerPointEntity extract(String inFilePath)
			throws Exception {
		PowerPointEntity temp = null;
		if(ExtensionUtils.isPPt(inFilePath)){
			InputStream inputStream = null;
			try {
				File pptFile = new File(inFilePath);
				inputStream = new FileInputStream(pptFile);
				POIFSFileSystem fs = new POIFSFileSystem(inputStream); 
				if(ExtensionUtils.isPPt(inFilePath)){
					temp = new PowerPointEntity();
					
					HSLFSlideShow ppt = new HSLFSlideShow(fs);
					
					HeadersFooters hf = ppt.getNotesHeadersFooters();
					temp.setHeaderText(hf.getHeaderText());
					temp.setFooterText(hf.getFooterText());
					temp.setDateTimeText(hf.getDateTimeText());
					temp.setDateTimeFormat(hf.getDateTimeFormat());
					/*
					SummaryInformation info = ppt.getSummaryInformation();
					temp.setTitle(info.getTitle());
					temp.setAuthor(info.getAuthor());
					temp.setKeywords(info.getKeywords());
					temp.setCreateDateTime(info.getCreateDateTime());
					temp.setRevNumber(info.getRevNumber());*/
					
				}
			} finally {
				if (inputStream != null) {
					try {
						inputStream.close();
					} catch (Exception e) {
						System.err.println(e);
					}
				}
			}
		}
		return temp;
	}

	

	

}
