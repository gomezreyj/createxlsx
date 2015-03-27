package com.se.uti;

import org.apache.poi.ss.usermodel.CellStyle;

public class TableInfo {
	
	public String columnType;
	public String columnName;
	public int fieldSize;
	CellStyle cellStyle = null;
	
	
	public int getFieldSize() {
		return fieldSize;
	}
	public void setFieldSize(int filedSize) {
		this.fieldSize = filedSize;
	}
	public String getColumnType() {
		return columnType;
	}
	public void setColumnType(String columnType) {
		this.columnType = columnType;
	}
	
	public String getColumnName() {
		return columnName;
	}
	public void setColumnName(String columnName) {
		this.columnName = columnName;
	}
	
	public CellStyle getCellStyle() {
		return cellStyle;
	}
	public void setCellStyle(CellStyle cellStyle) {
		this.cellStyle = cellStyle;
	}
	

}
