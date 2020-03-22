package org.wenrong;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.io.File;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;

/**
 * @author 郑文荣
 */
public class Main {

    public static void main(String[] args) throws IOException {
        HSSFWorkbook workbook = getWorkbook();
        HSSFSheet sheet = workbook.createSheet("测试");
        HSSFCellStyle row1_style = createCellStyle(workbook, (short) 10, XSSFFont.COLOR_NORMAL, (short) 200, "宋体", HorizontalAlignment.CENTER);

        // 设置样式
        HSSFCellStyle style = workbook.createCellStyle();
        // 设置样式
        style.setFillForegroundColor(IndexedColors.SKY_BLUE.index);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        // 生成一种字体
        HSSFFont font = workbook.createFont();
        // 设置字体
        font.setFontName("微软雅黑");
        // 设置字体大小
        font.setFontHeightInPoints((short) 12);
        // 在样式中引用这种字体
        style.setFont(font);

        ScCell[] cell = new ScCell[]{
                new ScCell(style,0,0,"项目名称",0,2,0,0),
                new ScCell(style,0,1,"招标编号",0,2,1,1),
                new ScCell(style,0,2,"招标单位",0,2,2,2),
                new ScCell(style,0,3,"项目类别",0,2,3,3),
                new ScCell(style,0,4,"项目区域",0,2,4,4),
                new ScCell(style,0,5,"开标时间",0,2,5,5),
                new ScCell(style,0,6,"评标时间",0,2,6,6),
                new ScCell(style,0,7,"审核状态",0,2,7,7),
                new ScCell(style,0,8,"招标方式",0,2,8,8),
                new ScCell(style,0,9,"招标组织形式",0,2,9,9),
                new ScCell(style,0,10,"招标代理机构",0,2,10,10),
                new ScCell(style,0,11,"出席开标人员姓名(代理机构)",0,2,11,11),
                new ScCell(style,0,12,"缴款单位(中标单位)",0,2,12,12),
                new ScCell(style,0,13,"中标金额(元)",0,2,13,13),
                new ScCell(style,0,14,"中标金额说明",0,2,14,14),
                new ScCell(style,0,15,"缴款通知书时间",0,2,15,15),
                new ScCell(style,0,16,"缴款金额(元)",0,0,16,20),
                new ScCell(style,0,21,"换票情况",0,0,21,22),
                new ScCell(style,1,16,"合计5=(1+2)",1,2,16,16),
                new ScCell(style,1,17,"场地租赁费(1)",1,2,17,17),
                new ScCell(style,1,18,"服务费",1,1,18,20),
                new ScCell(style,1,21,"时间",1,2,21,21),
                new ScCell(style,1,22,"发票号码",1,2,22,22),
                new ScCell(style,2,18,"新标准(2)"),
                new ScCell(style,2,19,"旧标准(2)"),
                new ScCell(style,2,20,"减负情况4=(3-2)")
        };

        generateTableHeader(sheet, cell);


        File file = new File("C://Users//Administrator//Desktop//测试1.xls");
        workbook.write(file);
    }

    /**
     * 生成表格头
     * @param sheet
     * @param cell
     */
    private static void generateTableHeader(HSSFSheet sheet, ScCell[] cell) {
        for(int i = 0;i < cell.length;i++){

            ScCell scCell = cell[i];
            HSSFRow hssfRow = sheet.getRow(scCell.getRow());
            if(hssfRow == null){
                hssfRow = sheet.createRow(scCell.getRow());
            }

            HSSFCell hssfCell = hssfRow.getCell(scCell.getColumn());
            if(hssfCell == null){
                hssfCell = hssfRow.createCell(scCell.getColumn());
            }

            //合并单元格
            if(scCell.isMergerCell()){

                for(int j = scCell.getMergerRowStart();j <= scCell.getMergerRowEnd();j++){

                    for(int k =  scCell.getMergerColumnStart(); k <= scCell.getMergerColumnEnd();k++){

                        HSSFRow tempRow = sheet.getRow(j);
                        if(tempRow == null){
                            tempRow = sheet.createRow(j);
                        }

                        HSSFCell tempCell = tempRow.getCell(k);
                        if(tempCell == null){
                            tempCell = tempRow.createCell(k);
                        }
                        tempCell.setCellStyle(scCell.getHssfCellStyle());

                    }

                }

                sheet.addMergedRegion(new CellRangeAddress(scCell.getMergerRowStart(),scCell.getMergerRowEnd(),scCell.getMergerColumnStart(),scCell.getMergerColumnEnd()));
            }


            //设置单元格的值
            Object cellValue = scCell.getValue();
            if(cellValue instanceof Boolean){
                hssfCell.setCellValue((Boolean) cellValue);
            } else if(cellValue instanceof Date){
                hssfCell.setCellValue((Date) cellValue);
            } else if (cellValue instanceof String){
                hssfCell.setCellValue((String) cellValue);
            } else if (cellValue instanceof Double){
                hssfCell.setCellValue((Double)cellValue );
            }else if (cellValue instanceof Calendar){
                hssfCell.setCellValue((Calendar)cellValue);
            }else if (cellValue instanceof RichTextString){
                hssfCell.setCellValue((RichTextString)cellValue);
            }

            //设置单元格的样式
            if(scCell.getHssfCellStyle() != null){
                hssfCell.setCellStyle(scCell.getHssfCellStyle());
            }

            sheet.autoSizeColumn(scCell.getColumn(), true);// 根据字段长度自动调整列的宽度


        }
    }

    public static void setCellBorder(HSSFWorkbook workbook){
        HSSFCellStyle cellStyle= workbook.createCellStyle();



    }

    private static HSSFCellStyle createCellStyle(HSSFWorkbook workbook,short fontHeightInPoints,short color,short fontHeight,String fontName,HorizontalAlignment align) {
        HSSFFont font = workbook.createFont();
        // 表头字体大小
        if(fontHeightInPoints != 0){

            //font.setFontHeightInPoints((short) 6);
            font.setFontHeightInPoints(fontHeightInPoints);

        }



        //字体颜色
        if(color != 0){

            //font.setColor(XSSFFont.COLOR_NORMAL);
            font.setColor(color);
        }


        if(fontHeight != 0){
           // font.setFontHeight((short) 200);
           font.setFontHeight(fontHeight);
        }


        // 表头字体名称
        //font.setFontName("宋体");
        if(null != fontName && !"".equals(fontName)){
            font.setFontName(fontName);
        }

        HSSFCellStyle cellStyle = workbook.createCellStyle();
        if(align != null){
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
        }

        cellStyle.setFont(font);
        return cellStyle;
    }

    /**
     * 创建一个row
     * @param sheet
     * @param rowNum
     * @param rowHeight
     * @param cellStyle
     * @return
     */
    public static HSSFRow createAndSetRowProperties(HSSFSheet sheet, int rowNum,float rowHeight, CellStyle cellStyle){
        HSSFRow row = sheet.createRow(rowNum);

        if(cellStyle != null){
            row.setRowStyle(cellStyle);
        }

        if(rowHeight != 0){
            row.setHeightInPoints(rowHeight);
        }

        return row;
    }

    /**
     * 创建一个单元格
     * @param row
     * @param column
     * @param cellValue
     * @param cellType
     * @param cellStyle
     * @return
     */
    public static HSSFCell createAndSetCellProperties(HSSFRow row, int column, Object cellValue, CellType cellType, CellStyle cellStyle){

        HSSFCell cell = row.createCell(column);

        if(cellStyle != null){
            cell.setCellStyle(cellStyle);
        }

        if(cellType != null){
            cell.setCellType(cellType);
        }

        if(cellValue instanceof Boolean){
            cell.setCellValue((Boolean) cellValue);
        } else if(cellValue instanceof Date){
            cell.setCellValue((Date) cellValue);
        } else if (cellValue instanceof String){
            cell.setCellValue((String) cellValue);
        } else if (cellValue instanceof Double){
            cell.setCellValue((Double)cellValue );
        }else if (cellValue instanceof Calendar){
            cell.setCellValue((Calendar)cellValue);
        }else if (cellValue instanceof RichTextString){
            cell.setCellValue((RichTextString)cellValue);
        }




        return cell;
    }


    public static HSSFWorkbook getWorkbook(){
        return new HSSFWorkbook();
    }

    /**
     * 对Cell单元格进行封装
     */
    static class ScCell {

        private HSSFCell hssfCell;
        private HSSFCellStyle hssfCellStyle;
        private Integer row;//行
        private Integer column;//列
        private Object value;

        public HSSFCellStyle getHssfCellStyle() {
            return hssfCellStyle;
        }

        public void setHssfCellStyle(HSSFCellStyle hssfCellStyle) {
            this.hssfCellStyle = hssfCellStyle;
        }

        /**
         * 合并单元格，行开始
         */
        private Integer mergerRowStart;

        /**
         * 合并单元格，行结束
         */
        private Integer mergerRowEnd;

        /**
         * 合并单元格，列开始
         */
        private Integer mergerColumnStart;

        /**
         * 合并单元格，列结束
         */
        private Integer mergerColumnEnd;


        public ScCell(HSSFCellStyle hssfCellStyle, Integer row, Integer column, Object value) {
            this.hssfCellStyle = hssfCellStyle;
            this.row = row;
            this.column = column;
            this.value = value;
        }

        public ScCell(HSSFCellStyle hssfCellStyle,Integer row, Integer column, Object value, Integer mergerRowStart, Integer mergerRowEnd, Integer mergerColumnStart, Integer mergerColumnEnd) {
            this.hssfCellStyle = hssfCellStyle;
            this.row = row;
            this.column = column;
            this.value = value;
            this.mergerRowStart = mergerRowStart;
            this.mergerRowEnd = mergerRowEnd;
            this.mergerColumnStart = mergerColumnStart;
            this.mergerColumnEnd = mergerColumnEnd;
        }

        public Object getValue() {
            return value;
        }

        public void setValue(Object value) {
            this.value = value;
        }

        public ScCell() {
        }

        public Integer getRow() {
            return row;
        }

        public void setRow(Integer row) {
            this.row = row;
        }

        public Integer getColumn() {
            return column;
        }

        public void setColumn(Integer column) {
            this.column = column;
        }

        public Integer getMergerRowStart() {
            return mergerRowStart;
        }

        public void setMergerRowStart(Integer mergerRowStart) {
            this.mergerRowStart = mergerRowStart;
        }

        public Integer getMergerRowEnd() {
            return mergerRowEnd;
        }

        public void setMergerRowEnd(Integer mergerRowEnd) {
            this.mergerRowEnd = mergerRowEnd;
        }

        public Integer getMergerColumnStart() {
            return mergerColumnStart;
        }

        public void setMergerColumnStart(Integer mergerColumnStart) {
            this.mergerColumnStart = mergerColumnStart;
        }

        public Integer getMergerColumnEnd() {
            return mergerColumnEnd;
        }

        public void setMergerColumnEnd(Integer mergerColumnEnd) {
            this.mergerColumnEnd = mergerColumnEnd;
        }

        public boolean isMergerCell(){
            if(mergerRowStart == null || mergerRowEnd == null
                    || mergerColumnStart == null || mergerColumnEnd == null){

                return false;
            }else {
                return true;
            }

        }
    }



}
