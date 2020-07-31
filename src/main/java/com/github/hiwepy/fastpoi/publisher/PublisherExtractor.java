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
package com.github.hiwepy.fastpoi.publisher;

import java.io.File;
import java.io.IOException;

import org.apache.poi.hpbf.extractor.PublisherTextExtractor;
import org.apache.poi.hpsf.SummaryInformation;

import com.github.hiwepy.fastpoi.publisher.entity.PublisherEntity;
import com.github.hiwepy.fastpoi.utils.ExtensionUtils;

public class PublisherExtractor   {

	private static PublisherExtractor instance = null;
	private static Object initLock = new Object();
	
	private PublisherExtractor(){};
	public static PublisherExtractor getInstance(){
		if (instance == null) {
			synchronized(initLock) {
				instance= new PublisherExtractor();
			}
		}
		return  instance;
	}
	
	public PublisherEntity extract(String inFilePath) throws IOException{
		return this.extract(new File(inFilePath));
	}
	
	public PublisherEntity extract(File inFilePath) throws IOException {
		PublisherEntity temp = null;
		if(ExtensionUtils.isPublisher(inFilePath.getAbsolutePath())){
			PublisherTextExtractor extactor = PublisherReader.getInstance().getTextExtactor(inFilePath);
			SummaryInformation info = extactor.getSummaryInformation();
			temp = new PublisherEntity();
			temp.setTitle(info.getTitle());
			temp.setAuthor(info.getAuthor());
			temp.setKeywords(info.getKeywords());
			temp.setText(extactor.getText().trim());
			temp.setCreateDateTime(info.getCreateDateTime());
			temp.setRevNumber(info.getRevNumber());
		}
		return temp;
	}

	

}
