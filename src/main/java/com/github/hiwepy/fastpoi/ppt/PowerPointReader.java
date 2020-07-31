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
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.sl.usermodel.SlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlideShow;

import com.github.hiwepy.fastpoi.utils.ExtensionUtils;

public class PowerPointReader {

	private static PowerPointReader instance = null;
	private static Object initLock = new Object();
	
	private PowerPointReader(){};
	public static PowerPointReader getInstance(){
		if (instance == null) {
			synchronized(initLock) {
				instance= new PowerPointReader();
			}
		}
		return  instance;
	}
	
	public HSLFSlideShow getHSLFSlideShow(String inFilePath) throws IOException {
		HSLFSlideShow hslf = null;
		if(ExtensionUtils.isPPt(inFilePath)){
			InputStream inputStream = null;
			try {
				File pptFile = new File(inFilePath);
				inputStream = new FileInputStream(pptFile);
				POIFSFileSystem fs = new POIFSFileSystem(inputStream); 
				if(ExtensionUtils.isPPt(inFilePath)){
					hslf = new HSLFSlideShow(fs);
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
		return hslf;
	}
	
	public SlideShow getSlideShow(String inFilePath) throws Exception {
		HSLFSlideShow hslf = getHSLFSlideShow(inFilePath);
		SlideShow ppt = null;/*
		if(hslf!=null){
			ppt = new SlideShow(hslf);
		}*/
		return ppt;
	}
	

	public XSLFSlideShow getXSLFSlideShow(String inFilePath) throws Exception {
		XSLFSlideShow xslf = null;
		if(ExtensionUtils.isPPtx(inFilePath)){
			xslf = new XSLFSlideShow(inFilePath);
		}
		return xslf;
	}
	

}
