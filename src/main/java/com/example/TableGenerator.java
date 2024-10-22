package com.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;

import javax.imageio.ImageIO;
import java.awt.Color;
import java.awt.Font;
import java.awt.Graphics2D;
import java.awt.FontMetrics;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.awt.FontFormatException;
import java.util.logging.Logger;
import java.util.logging.Level;
import java.util.ArrayList;
import java.util.List;

/**
 * TableGenerator 类用于生成表格并将其转换为图像。
 * 它使用 Apache POI 库来处理 Excel 文件，并使用 Java AWT 来生成图像。
 */
public class TableGenerator {
    private static final Logger LOGGER = Logger.getLogger(TableGenerator.class.getName());
    private Workbook workbook;
    private Sheet sheet;

    /**
     * 构造函数，初始化一个新的 Excel 工作簿和工作表。
     */
    public TableGenerator() {
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("Table");
    }

    /**
     * 向表格中添加一行数据。
     * @param data 要添加的数据数组
     */
    public void addRow(String[] data) {
        int lastRowNum = sheet.getLastRowNum();
        Row row = sheet.createRow(lastRowNum + 1);
        for (int i = 0; i < data.length; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(data[i]);
        }
    }

    /**
     * 将表格保存为 Excel 文件。
     * @param fileName 要保存的文件名
     * @throws IOException 如果保存过程中发生 IO 错误
     */
    public void saveAsExcel(String fileName) throws IOException {
        try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
            workbook.write(fileOut);
        }
    }

    /**
     * 合并指定列的多行单元格。
     * @param firstRow 开始行
     * @param lastRow 结束行
     * @param column 要合并的列
     */
    public void mergeRows(int firstRow, int lastRow, int column) {
        sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, column, column));
    }

    /**
     * 生成表格的图像并保存为文件。
     * @param fileName 要保存的图像文件名
     * @throws IOException 如果生成或保存图像过程中发生 IO 错误
     */
    public void generateImage(String fileName) throws IOException {
        int rows = sheet.getLastRowNum() + 1;
        int cols = sheet.getRow(0).getLastCellNum();

        int cellWidth = 100;
        int minCellHeight = 30;

        // 计算每行的实际高度
        int[] rowHeights = calculateRowHeights(rows, cols, cellWidth, minCellHeight);

        // 计算总高度
        int totalHeight = 0;
        for (int height : rowHeights) {
            totalHeight += height;
        }

        int width = cols * cellWidth;

        BufferedImage image = new BufferedImage(width, totalHeight, BufferedImage.TYPE_INT_RGB);
        Graphics2D g2d = image.createGraphics();

        g2d.setColor(Color.WHITE);
        g2d.fillRect(0, 0, width, totalHeight);

        g2d.setColor(Color.BLACK);
        
        // 加载自定义字体
        Font font = loadCustomFont();
        g2d.setFont(font);

        // 绘制表格和文本
        drawTableAndText(g2d, rows, cols, cellWidth, rowHeights);

        g2d.dispose();
        ImageIO.write(image, "png", new File(fileName));
    }

    /**
     * 加载自定义字体。
     * @return 加载的字体，如果加载失败则返回默认字体
     */
    private Font loadCustomFont() {
        Font font = null;
        try (InputStream is = getClass().getResourceAsStream("/fonts/msyh.ttc")) {
            if (is == null) {
                LOGGER.severe("找不到字体文件：/fonts/msyh.ttc");
                throw new IOException("找不到字体文件");
            }
            font = Font.createFont(Font.TRUETYPE_FONT, is).deriveFont(12f);
        } catch (FontFormatException e) {
            LOGGER.log(Level.SEVERE, "字体格式错误", e);
        } catch (IOException e) {
            LOGGER.log(Level.SEVERE, "读取字体文件时发生IO错误", e);
        }

        if (font == null) {
            LOGGER.severe("无法加载字体，使用默认字体");
            font = new Font("SansSerif", Font.PLAIN, 12);
        }
        return font;
    }

    /**
     * 计算每行的实际高度。
     */
    private int[] calculateRowHeights(int rows, int cols, int cellWidth, int minCellHeight) {
        int[] rowHeights = new int[rows];
        Font font = loadCustomFont();
        BufferedImage dummyImage = new BufferedImage(1, 1, BufferedImage.TYPE_INT_RGB);
        Graphics2D g2d = dummyImage.createGraphics();
        g2d.setFont(font);
        FontMetrics metrics = g2d.getFontMetrics();

        for (int i = 0; i < rows; i++) {
            int maxHeight = minCellHeight;
            Row row = sheet.getRow(i);
            if (row != null) {
                for (int j = 0; j < cols; j++) {
                    Cell cell = row.getCell(j);
                    if (cell != null) {
                        String text = cell.toString();
                        List<String> lines = wrapText(text, cellWidth, metrics);
                        int textHeight = lines.size() * metrics.getHeight();
                        maxHeight = Math.max(maxHeight, textHeight + 10); // 添加一些额外的空间
                    }
                }
            }
            rowHeights[i] = maxHeight;
        }
        g2d.dispose();
        return rowHeights;
    }

    /**
     * 绘制表格和文本。
     */
    private void drawTableAndText(Graphics2D g2d, int rows, int cols, int cellWidth, int[] rowHeights) {
        int y = 0;
        for (int i = 0; i < rows; i++) {
            Row row = sheet.getRow(i);
            int cellHeight = rowHeights[i];
            for (int j = 0; j < cols; j++) {
                int x = j * cellWidth;

                // 检查是否是合并单元格
                CellRangeAddress mergedRegion = getMergedRegion(i, j);
                if (mergedRegion != null) {
                    // 如果是合并单元格的左上角，绘制整个合并区域
                    if (i == mergedRegion.getFirstRow() && j == mergedRegion.getFirstColumn()) {
                        int mergedWidth = (mergedRegion.getLastColumn() - mergedRegion.getFirstColumn() + 1) * cellWidth;
                        int mergedHeight = 0;
                        for (int k = mergedRegion.getFirstRow(); k <= mergedRegion.getLastRow(); k++) {
                            mergedHeight += rowHeights[k];
                        }
                        g2d.drawRect(x, y, mergedWidth, mergedHeight);

                        Cell cell = row.getCell(j);
                        String cellValue = cell == null ? "" : cell.toString();
                        // 在合并单元格的中心绘制文本
                        drawCenteredString(g2d, cellValue, x, y, mergedWidth, mergedHeight);
                    }
                } else {
                    // 如果不是合并单元格，正常绘制
                    g2d.drawRect(x, y, cellWidth, cellHeight);
                    Cell cell = row.getCell(j);
                    String cellValue = cell == null ? "" : cell.toString();
                    drawCenteredString(g2d, cellValue, x, y, cellWidth, cellHeight);
                }
            }
            y += cellHeight;
        }
    }

    /**
     * 获取指定单元格所在的合并区域。
     * @param row 行索引
     * @param column 列索引
     * @return 合并区域，如果不是合并单元格则返回 null
     */
    private CellRangeAddress getMergedRegion(int row, int column) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            if (range.isInRange(row, column)) {
                return range;
            }
        }
        return null;
    }

    /**
     * 在指定区域内绘制居中的文本。
     * @param g2d Graphics2D 对象
     * @param text 要绘制的文本
     * @param x 区域左上角 x 坐标
     * @param y 区域左上角 y 坐标
     * @param width 区域宽度
     * @param height 区域高度
     */
    private void drawCenteredString(Graphics2D g2d, String text, int x, int y, int width, int height) {
        FontMetrics metrics = g2d.getFontMetrics(g2d.getFont());
        List<String> lines = wrapText(text, width, metrics);
        
        int lineHeight = metrics.getHeight();
        int totalTextHeight = lineHeight * lines.size();
        
        int startY = y + (height - totalTextHeight) / 2 + metrics.getAscent();
        
        for (String line : lines) {
            int lineWidth = metrics.stringWidth(line);
            int startX = x + (width - lineWidth) / 2;
            g2d.drawString(line, startX, startY);
            startY += lineHeight;
        }
    }

    /**
     * 将文本按指定宽度换行。
     * @param text 要换行的文本
     * @param width 可用宽度
     * @param metrics 字体度量
     * @return 换行后的文本行列表
     */
    private List<String> wrapText(String text, int width, FontMetrics metrics) {
        List<String> lines = new ArrayList<>();
        String[] words = text.split("\\s+");
        StringBuilder currentLine = new StringBuilder();

        for (String word : words) {
            if (currentLine.length() == 0) {
                currentLine.append(word);
            } else if (metrics.stringWidth(currentLine + " " + word) <= width) {
                currentLine.append(" ").append(word);
            } else {
                lines.add(currentLine.toString());
                currentLine = new StringBuilder(word);
            }
        }

        if (currentLine.length() > 0) {
            lines.add(currentLine.toString());
        }

        // 如果只有一行，但是超出宽度，则强制换行
        if (lines.size() == 1 && metrics.stringWidth(lines.get(0)) > width) {
            return forceSplitLine(lines.get(0), width, metrics);
        }

        return lines;
    }

    /**
     * 强制按字符分割文本行。
     * @param text 要分割的文本
     * @param width 可用宽度
     * @param metrics 字体度量
     * @return 分割后的文本行列表
     */
    private List<String> forceSplitLine(String text, int width, FontMetrics metrics) {
        List<String> lines = new ArrayList<>();
        StringBuilder currentLine = new StringBuilder();

        for (char c : text.toCharArray()) {
            if (metrics.stringWidth(currentLine.toString() + c) <= width) {
                currentLine.append(c);
            } else {
                lines.add(currentLine.toString());
                currentLine = new StringBuilder(String.valueOf(c));
            }
        }

        if (currentLine.length() > 0) {
            lines.add(currentLine.toString());
        }

        return lines;
    }
}
