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

import java.sql.SQLException;
import java.util.List;

public class XxlsPrint extends XxlsAbstract {

	@Override
	public void optRows(int sheetIndex,int curRow, List<String> rowlist) throws SQLException {
		for (int i = 0; i < rowlist.size(); i++) {
			System.out.print("'" + rowlist.get(i) + "',");
		}
		System.out.println();
	}

	public static void main(String[] args) throws Exception {
		XxlsPrint howto = new XxlsPrint();
		howto.processOneSheet("D:/bb.xlsx",1);
//		howto.processAllSheets("F:/new.xlsx");
	}
}
