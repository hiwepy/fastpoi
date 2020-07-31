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
package com.github.hiwepy.fastpoi.utils;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hslf.usermodel.HSLFSlideShow;

public class PPtUtils {
	
	public static void writePPt(HSLFSlideShow hslf, String outFilePath) throws IOException {
		if(hslf!=null&&outFilePath!=null){
			FileOutputStream fos = null;
			try {
				File file = new File(outFilePath);
				fos = new FileOutputStream(file);
				hslf.write(fos);
			} finally {
				try {
					if (fos != null) {
						fos.close();
					}
				} catch (IOException e) {
					LogUtils.error(e);
				}
			}
		}
	}

}
