package com.example;

import java.io.IOException;

public class Main {
    public static void main(String[] args) {
        TableGenerator generator = new TableGenerator();

        // 添加表头
        generator.addRow(new String[]{"部门", "姓名", "年龄", "职业"});

        // 添加数据
        generator.addRow(new String[]{"技术部", "张三", "25", "工程师"});
        generator.addRow(new String[]{"", "李四死死死死四重垦局阿赛洛烦死撒娇发腮了;", "28", "设计师"});
        generator.addRow(new String[]{"", "王五", "87896876876575764565465453543564654535435643543245676530", "产品经理"});
        generator.addRow(new String[]{"市场部", "赵六", "32", "销售"});
        generator.addRow(new String[]{"", "钱七", "35", "市场经理"});

        // 合并单元格
        generator.mergeRows(1, 3, 0); // 合并技术部
        generator.mergeRows(4, 5, 0); // 合并市场部

        try {
            // 保存为Excel文件
            generator.saveAsExcel("table_merged.xlsx");
            System.out.println("Excel文件已生成: table_merged.xlsx");

            // 生成图片
            generator.generateImage("table_merged.png");
            System.out.println("图片已生成: table_merged.png");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
