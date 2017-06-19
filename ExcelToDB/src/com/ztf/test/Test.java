package com.ztf.test;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.util.List;
import java.util.Map;

import com.ztf.excelreader.ExcelReader;

public class Test {
	public static void main(String[] args) {
		try {
			String url = "jdbc:oracle:thin:@localhost:1521:orcl";
			Class.forName("oracle.jdbc.driver.OracleDriver");
			Connection conn = DriverManager.getConnection(url, "system", "system");
			// 关闭事务的自动提交
			conn.setAutoCommit(false);
			ExcelReader excelReader = new ExcelReader();
			List<Map<Integer, String>> contentList = excelReader.readExcelContent("c:\\test.xls");
			System.out.println("获得Excel表格的内容:" + contentList.size());
			// 获得Map的key的个数
			int cloumtCount = contentList.get(0).keySet().size();
			System.out.println(contentList.size());
			String sql = "insert into test_content values(?,?,?,?)";
			PreparedStatement pstmt = conn.prepareStatement(sql);
			for (int i = 1; i < contentList.size(); i++) {
				for (int j = 1; j <= cloumtCount; j++) {
					pstmt.setObject(j, contentList.get(i).get(j - 1));
				}
				pstmt.addBatch();
			}
			pstmt.executeBatch();
			conn.commit();
			pstmt.close();
			conn.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
