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

import org.apache.log4j.Logger;

public final class LogUtils {
	
	private static Logger log = Logger.getLogger(LogUtils.class);
	
	public static void error(Exception e){
		if(log.isTraceEnabled()){
			log.error(new StringBuffer("[").append(e.getMessage()).append("]").append(e.getCause()));
		}
	}
	
	public static void debug(Exception e){
		if(log.isDebugEnabled()){
			log.debug(new StringBuffer("[").append(e.getMessage()).append("]").append(e.getCause()));
		}
		log.debug(new StringBuffer("[").append(e.getMessage()).append("]").append(e.getCause()));
	}
	
	public static void info(Exception e){
		if(log.isInfoEnabled()){
			log.info(new StringBuffer("[").append(e.getMessage()).append("]").append(e.getCause()));
		}
	}
}
