package com.lxm.excel.parser;

import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;

import org.apache.commons.lang.StringUtils;

import com.lxm.excel.parser.commerce.DataGatherer;



public class Main {

    public static void main(String[] args) throws IOException {
    	System.out.println("请输入提取数据目录：");
    	BufferedReader reader = new BufferedReader(new InputStreamReader(System.in));
    	String dir = reader.readLine();
    	if(StringUtils.isBlank(dir)){
    		dir = "./";
    	}
    	File dirFile = new File(dir);
    	if(!dirFile.isDirectory()){
    		throw new RuntimeException(dir + " 为非合法目录");
    	}
    	String targetFile = new File(dirFile, "流向填报表.xlsx").getAbsolutePath();
    	String dataDir = null;
    	String templateFile = null;
    	for(File file: dirFile.listFiles()){
    		if(file.getName().startsWith("原始数据")){
    			dataDir = file.getAbsolutePath();
    		}else if(file.getName().startsWith("解析模板.xlsx")){
    			templateFile = file.getAbsolutePath();
    		}
    	}
    	if(templateFile == null){
    		throw new RuntimeException(dir + "目录下未找到解析模板文件：解析模板.xlsx");
    	}
    	if(dataDir == null){
    		throw new RuntimeException(dir + "目录下未找到解析数据目录：原始目录");
    	}
        DataGatherer dataGatherer = new DataGatherer(templateFile, dataDir, targetFile);
    	dataGatherer.extractData();
    }

}
