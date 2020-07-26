package com.lxm.excel.parser.commerce;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang.StringUtils;

import com.lxm.excel.parser.model.GatherModel;
import com.lxm.excel.parser.utils.PoiExcelUtil;

/**
 * 
 * ClassName: com.lxm.excel.parser.commerce.DataGatherer <br/>
 * Function: 商业公司数据收集器 <br/>
 * Date: 2020年7月25日 下午9:12:53 <br/>
 * @author liuxiangming <br/>
 */
public class DataGatherer {
	
	private Map<String, GatherModel> gatherModelMap = new HashMap<>();
	
	private String[] gatherTitle;
	
	private String targetFile;
	
	private String dataDir;
	
	/**
	 * 
	 * @param templateFile 解析模板文件路径
	 * @param dataDir 原始数据目录
	 * @param targetFile 结果文件路径
	 */
	public DataGatherer(String templateFile, String dataDir, String targetFile){
		this.dataDir = dataDir;
		init(templateFile, targetFile);
	}
	
	/**
	 * 提取汇总数据
	 */
	public void extractData(){
		File dir = new File(dataDir);
		if(!dir.isDirectory()){
			throw new RuntimeException(dataDir + "不是有效数据目录");
		}
		List<String[]> resultData = new ArrayList<>();
		resultData.add(gatherTitle);
		File[] files = dir.listFiles();
		for(File file: files){
			extractFile(file, resultData);
		}
		File target = new File(targetFile);
		System.out.println("导出结果文件：" + targetFile);
		PoiExcelUtil.createExcel(target, "汇总结果");
		PoiExcelUtil.exportExcel(target, resultData, 0, 0);
	}
	
	private void extractFile(File file, List<String[]> resultData){
		String fileName = file.getName();
		if(!fileName.toLowerCase().endsWith(".xls") && !fileName.toLowerCase().endsWith(".xlsx")){ //不是有效excel文件
			return;
		}
		System.out.println("解析文件：" + file.getAbsolutePath());
		String commerceName = fileName.substring(0, fileName.lastIndexOf("."));
		GatherModel gatherModel = gatherModelMap.get(commerceName.trim());
		if(gatherModel == null){
			throw new RuntimeException(fileName + "不存在有效解析模板数据");
		}
		List<String[]> datas = PoiExcelUtil.parseExcel(file, 0, gatherModel.getStartRow() - 1, 0);
		for(String[] dataArr: datas){
			Integer[] columnIndexes = gatherModel.getColumnIndexes();
			String[] targetArr = new String[columnIndexes.length + 1];
			targetArr[0] = commerceName;
			for(int i = 0; i < columnIndexes.length; i++){
				if(columnIndexes[i] == null){
					targetArr[i + 1] = "";
					continue;
				}
				if(dataArr.length <= columnIndexes[i]){
					throw new RuntimeException(fileName + "解析超出数据范围");
				}
				targetArr[i + 1] = dataArr[columnIndexes[i]];
			}
			resultData.add(targetArr);
		}
	}
	
	/**
	 * 
	 * @param templateFile
	 * @param dataDir
	 * @param targetFile
	 */
	private void init(String templateFile, String targetFile){
		this.targetFile = targetFile;
		File template = new File(templateFile);
		List<String[]> datas = PoiExcelUtil.parseExcel(template, 0, 0, 0);
		String[] titleArr = datas.get(0); // 实际配送商业	数据开始行	供货方	流向时间	终端名称	药品名	规格	数量单位	数量	批号
		List<String> titleList = new ArrayList<>();
		for(int i = 0; i < titleArr.length; i++){
			if(i == 1){ //数据开始行，跳过
				continue;
			}
			titleList.add(titleArr[i].trim());
		}
		gatherTitle = titleList.toArray(new String[0]); // 实际配送商业	供货方	流向时间	终端名称	药品名	规格	数量单位	数量	批号
		for(int i = 1; i < datas.size(); i++){
			String[] dataArr = datas.get(i);
			Integer[] columnIndexes = new Integer[dataArr.length - 2];//供货方 流向时间 终端名称 药品名 规格 数量单位 数量 批号
			for(int j = 2; j < dataArr.length; j++){
				if(StringUtils.isBlank(dataArr[j])){
					columnIndexes[j - 2] = null;
				}else{
					columnIndexes[j - 2] = PoiExcelUtil.column2Index(dataArr[j].trim());
				}
			}
			gatherModelMap.put(dataArr[0].trim(), new GatherModel(Integer.valueOf(dataArr[1].trim()), columnIndexes));
		}
	}

}
