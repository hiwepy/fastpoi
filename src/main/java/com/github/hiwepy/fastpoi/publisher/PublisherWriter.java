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
import java.io.FileWriter;
import java.io.IOException;

import org.apache.poi.hpbf.extractor.PublisherTextExtractor;

import com.github.hiwepy.fastpoi.utils.LogUtils;

public class PublisherWriter   {

	private static PublisherWriter instance = null;
	private static Object initLock = new Object();
	
	private PublisherWriter(){};
	public static PublisherWriter getInstance(){
		if (instance == null) {
			synchronized(initLock) {
				instance= new PublisherWriter();
			}
		}
		return  instance;
	}
	
	
	public void writeToLocal(PublisherTextExtractor extactor,String outFilePath) throws  IOException{
		this.writeToLocal(extactor, new File(outFilePath));
	}
	
	public void writeToLocal(PublisherTextExtractor extactor,File outFile) throws  IOException{
		FileWriter writer = null;
		try {
			writer = new FileWriter(outFile);
			writer.write(extactor.getText().trim());
		} finally {
			if(writer != null) {
				try {
					writer.close();
				} catch (Exception e) {
					LogUtils.error(e);
				}
			}
		}
	}

}
