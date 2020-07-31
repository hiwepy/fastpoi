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
package com.github.hiwepy.fastpoi.visio;

import org.apache.poi.hdgf.extractor.VisioTextExtractor;
import org.apache.poi.hpsf.SummaryInformation;

import com.github.hiwepy.fastpoi.utils.ExtensionUtils;
import com.github.hiwepy.fastpoi.visio.entity.VisioEntity;

public class VisioExtractor{

	private static VisioExtractor instance = null;
	private static Object initLock = new Object();
	
	private VisioExtractor(){};
	public static VisioExtractor getInstance(){
		if (instance == null) {
			synchronized(initLock) {
				instance= new VisioExtractor();
			}
		}
		return  instance;
	}

	public VisioEntity extract(String inFilePath) throws Exception {
		VisioEntity temp = null;
		if(ExtensionUtils.isVisio(inFilePath)){
			VisioTextExtractor extactor = VisioReader.getInstance().getTextExtactor(inFilePath);
			SummaryInformation info = extactor.getSummaryInformation();
			temp = new VisioEntity();
			temp.setTitle(info.getTitle());
			temp.setAuthor(info.getAuthor());
			temp.setKeywords(info.getKeywords());
			temp.setCreateDateTime(info.getCreateDateTime());
			temp.setRevNumber(info.getRevNumber());
			temp.setText(extactor.getText());
			temp.setAllTtext(extactor.getAllText());
		}
		return temp;
	}


}
