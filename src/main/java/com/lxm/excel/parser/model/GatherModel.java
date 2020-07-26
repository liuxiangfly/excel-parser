package com.lxm.excel.parser.model;

/**
 * 
 * ClassName: com.lxm.excel.parser.model.GatherModel <br/>
 * Function: 商业公司数据收集模型<br/>
 * Date: 2020年7月25日 下午9:07:36 <br/>
 * @author liuxiangming <br/>
 */
public class GatherModel {
	
	/**
	 * 商业公司原始数据开始行
	 */
	private Integer startRow;
	
	/**
	 * 商业公司原始数据收集列(index从0开始)
	 */
	private Integer[] columnIndexes;
	

	public GatherModel(Integer startRow, Integer[] columnIndexes) {
		super();
		this.startRow = startRow;
		this.columnIndexes = columnIndexes;
	}

	public Integer[] getColumnIndexes() {
		return columnIndexes;
	}

	public void setColumnIndexes(Integer[] columnIndexes) {
		this.columnIndexes = columnIndexes;
	}

	public Integer getStartRow() {
		return startRow;
	}

	public void setStartRow(Integer startRow) {
		this.startRow = startRow;
	}
	
}
