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

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.imageio.ImageIO;

import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.sl.usermodel.Slide;
import org.apache.poi.sl.usermodel.SlideShow;

import com.github.hiwepy.fastpoi.utils.LogUtils;

public class PowerPointWriter{

	private static PowerPointWriter instance = null;
	private static Object initLock = new Object();
	
	private PowerPointWriter(){};
	public static PowerPointWriter getInstance(){
		if (instance == null) {
			synchronized(initLock) {
				instance= new PowerPointWriter();
			}
		}
		return  instance;
	}
	
	public void writeImageToLocal(String inFilePath,String outFileDir,String suffix) throws IOException {
		HSLFSlideShow show = PowerPointReader.getInstance().getHSLFSlideShow(inFilePath);
		this.writeImageToLocal(show, outFileDir, new File(inFilePath).getName(), suffix);
	}
	
	public void writeImageToLocal(String inFilePath,File outFileDir,String suffix) throws IOException {
		HSLFSlideShow show = PowerPointReader.getInstance().getHSLFSlideShow(inFilePath);
		this.writeImageToLocal(show, outFileDir, new File(inFilePath).getName(), suffix);
	}
	
	public void writeImageToLocal(HSLFSlideShow show,String outFileDir,String preffix,String suffix) throws IOException {
		this.writeImageToLocal(show, new File(outFileDir), preffix, suffix);
	}
	
	public void writeImageToLocal(HSLFSlideShow show,File outFileDir,String preffix,String suffix) throws IOException {
		//SlideShow ppt = PowerPointReader.getInstance().getSlideShow(show);  
		//this.writeImageToLocal(ppt, outFileDir, preffix, suffix);
	}
	
	public void writeImageToLocal(SlideShow show,File outFileDir,String preffix,String suffix) throws IOException {
		//建立图片文件目录
		Dimension pgsize = show.getPageSize();  
		Slide[] slide = null;//show.getSlides();  
		for (int index = 0; index < slide.length; index++) {  
		    BufferedImage img = new BufferedImage(pgsize.width, pgsize.height , BufferedImage.TYPE_INT_RGB);  
		    Graphics2D graphics = img.createGraphics();  
		    //clear the drawing area  
		    graphics.setPaint(Color.white);  
		    graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width, pgsize.height));  
		    //render  
		    slide[index].draw(graphics);  
		    FileOutputStream fo = null;
        	try {
				fo = new FileOutputStream(new File(outFileDir, preffix+index+suffix));
        		ImageIO.write(img, suffix.substring(1), fo); 
        	} finally {
				if(fo != null) {
					try {
						fo.close();
					} catch (Exception e) {
						LogUtils.error(e);
					}
				}
			}
		}
	}
	

}
