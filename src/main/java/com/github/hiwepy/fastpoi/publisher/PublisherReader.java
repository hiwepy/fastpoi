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
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hpbf.extractor.PublisherTextExtractor;

import com.github.hiwepy.fastpoi.utils.ExtensionUtils;
import com.github.hiwepy.fastpoi.utils.LogUtils;

public class PublisherReader   {

	private static PublisherReader instance = null;
	private static Object initLock = new Object();
	
	private PublisherReader(){};
	public static PublisherReader getInstance(){
		if (instance == null) {
			synchronized(initLock) {
				instance= new PublisherReader();
			}
		}
		return  instance;
	}
	
	public PublisherTextExtractor getTextExtactor(File file) throws IOException{
		if(file.exists()){
			if(ExtensionUtils.isPublisher(file.getAbsolutePath())){
				PublisherTextExtractor extactor  = null;
				InputStream inputStream = null;
				try {
					inputStream = new FileInputStream(file);
					extactor = new PublisherTextExtractor(inputStream);
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
	
	public PublisherTextExtractor getTextExtactor(String inFilePath) throws IOException{
		return this.getTextExtactor(new File(inFilePath) );
	}
	

}
