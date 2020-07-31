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

import java.beans.BeanInfo;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.util.regex.Pattern;

import org.apache.log4j.Logger;

public class ExcelReflectUtil{

	/**
	 * apache.commons.logging  组件
	 */
	public static Logger logr = Logger.getLogger(ExcelReflectUtil.class);
	public static final String regexp = "^\\$\\(.*?\\)$";

	public static boolean isMatche(String source) {
		return Pattern.matches(regexp, source);
	}
	
	public static String getFlagName(String source) {
		if(source!=null&&isMatche(source)){
			return source.substring("$(".length()-1, source.length()-1);
		}
		return null;
	}
	
	public static <R> Object reflectValueFromBean(String propertyName, R object){
		try {
			BeanInfo beanInfo = null;
			Object result = null;
			// 获取类属性
			beanInfo = Introspector.getBeanInfo(object.getClass());
			for (PropertyDescriptor descriptor : beanInfo.getPropertyDescriptors()) {
				String propertyName2 = descriptor.getName();
				if (propertyName.equals(propertyName2)) {
					result = descriptor.getReadMethod().invoke(object);
					break;
				}
			}
			return result;
		} catch (Exception e1) {
			return null;
		}
	}

	/*public static <R> Object reflectValueFromBean2(String propertyName, R object){
		try {
			Object result = null;
			Field[] fields = object.getClass().getDeclaredFields();
			for (Field field : fields) {
				String key = field.getName();
				//4.如果匹配，提取值，进行设置
				if(propertyName.equals(key)){
					result = BeanUtils.getProperty(object, key);
				}
			}
			return result;
		} catch (Exception e1) {
			return null;
		}
	}
	
	public static <R> Object reflectValueFromBean3(String propertyName, R object){
		try {
			Object result = null;
			PropertyDescriptor[] propertys = Introspector.getBeanInfo(object.getClass()).getPropertyDescriptors();
			for (PropertyDescriptor property : propertys) {
				String key = property.getName();
				//4.如果匹配，提取值，进行设置
				if(key.equals(propertyName)){
					result = BeanUtils.getProperty(object, key);
				}
			}
			return result;
		} catch (Exception e1) {
			return null;
		}
	}*/
	
}
