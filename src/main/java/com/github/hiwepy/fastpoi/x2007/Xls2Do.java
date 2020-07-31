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
package com.github.hiwepy.fastpoi.x2007;

import java.util.Date;

import com.github.hiwepy.fastpoi.x2003.HxlsBig;
import com.github.hiwepy.fastpoi.x2003.HxlsPrint;

public class Xls2Do {

	public static void main(String[] args) {
		String filename = "E:/up.xls";
		String hxlsTable = "score_temp";
		String xlsxTable = "temp_student";
		Date date1 = new Date();
		System.out.println(date1.toString());
//		xls2Database(filename,hxlsTable,xlsxTable);
		Date date2 = new Date();
		System.out.println(date2.toString());
		xls2Print(filename);
	}
	
//	public static void main(String[] args){
//		String filename = "E:/09bu.xlsx";
//		String xlsxTable = "score_temp";
//		XxlsBig howto;
//		try {
//			howto = new XxlsBig(xlsxTable);
//			howto.setCreate(true);
//			howto.process(filename);
//			howto.close();
//		} catch (Exception e) {
//			e.printStackTrace();
//		}
//		
////		filename = "E:/2009new.xlsx";
////		xlsxTable = "early_data";
////		try {
////			howto = new XxlsBig(xlsxTable);
////			howto.setCreate(true);
////			howto.process(filename);
////			howto.close();
////		} catch (Exception e) {
////			e.printStackTrace();
////		}
//		
//		filename = "E:/2009old.xlsx";
//		xlsxTable = "early_data";
//		try {
//			howto = new XxlsBig(xlsxTable);
//			howto.setCreate(true);
//			howto.process(filename);
//			howto.close();
//		} catch (Exception e) {
//			e.printStackTrace();
//		}
//	}
	
	//excel->databse
	public static void xls2Database(String filename,String hxlsTable,String xlsxTable){

		String fileType = filename.substring(filename.lastIndexOf(".")+1).toLowerCase();
		try {
			if (fileType.equals("xls")) {
				HxlsBig xls2csv = new HxlsBig(filename, hxlsTable);
				xls2csv.process();
				xls2csv.close();
			} else if (fileType.equals("xlsx")) {
				XxlsBig howto = new XxlsBig(xlsxTable);
				howto.setCreate(true);
				// howto.processOneSheet("F:/new.xlsx",1);
				howto.process(filename);
				howto.close();
			}
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}	
	}
	
	//excel->print
	public static void xls2Print(String filename){
		String fileType = filename.substring(filename.lastIndexOf(".")+1).toLowerCase();
		try {
			if (fileType.equals("xls")) {
				HxlsPrint xls2csv;
				xls2csv = new HxlsPrint(filename);
				xls2csv.process();
			} else if (fileType.equals("xlsx")) {
				XxlsPrint howto = new XxlsPrint();
				howto.setTitleRow(0);
				howto.process(filename);
			}
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}
}
