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
import java.io.OutputStream;

import org.apache.poi.hsmf.datatypes.AttachmentChunks;

public class OutlookUtils {
	
	/**
	 * Processes a single attachment: reads it from the Outlook MSG file and
	 * writes it to disk as an individual file.
	 *
	 * @param attachment the chunk group describing the attachment
	 * @param dir the directory in which to write the attachment file
	 * @throws IOException when any of the file operations fails
	 */
	public static void processAttachment(AttachmentChunks attachment,File dir) throws IOException {
	   String fileName = attachment.attachFileName.toString();
	   if(attachment.attachLongFileName != null) {
	      fileName = attachment.attachLongFileName.toString();
	   }
		File f = new File(dir, fileName);
		OutputStream fileOut = null;
		try {
			fileOut = new FileOutputStream(f);
			fileOut.write(attachment.attachData.getValue());
		} finally {
			if(fileOut != null) {
				fileOut.close();
			}
		}
	}
	
}
