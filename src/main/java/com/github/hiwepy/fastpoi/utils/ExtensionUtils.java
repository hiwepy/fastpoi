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

import org.apache.commons.io.FilenameUtils;

import com.github.hiwepy.fastxls.core.utils.Assert;
/**
 * 文档后缀判断方法
 */
public class ExtensionUtils {
	
	public static boolean isXls(String filePath) {
		return ExtensionUtils.isXls(new File(filePath));
	}
	
	public static boolean isXls(File file) {
		Assert.notNull(file, " file is not specified!");
		Assert.isTrue(file.exists(), " file is not found !");
		Assert.isTrue(file.isFile(), " file is not a file !");
		return "xls".equalsIgnoreCase(FilenameUtils.getExtension(file.getName()));
	}
	
	public static boolean isXlsx(String filePath) {
		return ExtensionUtils.isXlsx(new File(filePath));
	}
	
	public static boolean isXlsx(File file) {
		Assert.notNull(file, " file is not specified!");
		Assert.isTrue(file.exists(), " file is not found !");
		Assert.isTrue(file.isFile(), " file is not a file !");
		return "xlsx".equalsIgnoreCase(FilenameUtils.getExtension(file.getName()));
	}

	public static boolean isVisio(String inFilePath) {
		// TODO Auto-generated method stub
		return false;
	}

	public static boolean isPublisher(String absolutePath) {
		// TODO Auto-generated method stub
		return false;
	}

	public static boolean isPPt(String inFilePath) {
		// TODO Auto-generated method stub
		return false;
	}

	public static boolean isPPtx(String inFilePath) {
		// TODO Auto-generated method stub
		return false;
	}

	public static boolean isOutlook(String inFilePath) {
		// TODO Auto-generated method stub
		return false;
	}
	
	
}
