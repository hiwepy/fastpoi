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

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;

import org.apache.poi.hdgf.extractor.VisioTextExtractor;

import com.github.hiwepy.fastpoi.utils.LogUtils;

public class VisioWriter {

	private static VisioWriter instance = null;
	private static Object initLock = new Object();
	
	private VisioWriter(){};
	public static VisioWriter getInstance(){
		if (instance == null) {
			synchronized(initLock) {
				instance= new VisioWriter();
			}
		}
		return  instance;
	}
	
	public void visio2Text(String inFilePath, String outFilePath) throws IOException {
		VisioTextExtractor extactor = VisioReader.getInstance().getTextExtactor(inFilePath);
		FileWriter writer = null;
		try {
			writer = new FileWriter(new File(outFilePath));
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
