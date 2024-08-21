package org.example;

import cn.hutool.json.JSONArray;
import cn.hutool.json.JSONObject;
import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.net.URL;
import java.net.URLConnection;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;


/**
 * @Description
 * @Author 罗宇航
 * @Date 2024/3/28
 */
public class Main {
    private static File staticSourceImgFile;
    private static String staticTargetImgFilePath;
    private static String staticOutputFilePath;

    private static String tablePreName;
    private static String date;

    private static int inputTableIndex;

    private static File[] newImgFile;

    public static void main(String[] args) {
//        System.setProperty("file.encoding", "UTF-8");
        //解决路径包含中文的情况

        try {
            // 创建一个Scanner对象，用于从控制台读取输入
//            Scanner scanner = new Scanner(System.in);
//
//            System.out.print("输入第一个参数（示例：pu）：");
//            tablePreName = scanner.nextLine();
//            System.out.print("输入日期（示例：3.26）：");
//            date = scanner.nextLine();
//            System.out.print("输入起始表格序号：");
//            inputTableIndex = Integer.valueOf(scanner.nextLine());
//            System.out.print("输入源文件路径: ");
//            String inputSourceExcelPath = scanner.nextLine();
//            System.out.println("源文件路径============> "+inputSourceExcelPath);
//            System.out.print("输入输出文件夹路径: ");
//            staticOutputFilePath = scanner.nextLine();
//            System.out.print("输入源图库路径: ");
//            String inputSourceImgPath = scanner.nextLine();
//            System.out.print("输入目标图库文件夹路径: ");
//            staticTargetImgFilePath = scanner.nextLine();


        tablePreName = "pu";
        date = "4.1";
        inputTableIndex = 2;
        String inputSourceExcelPath = "/Users/fury/workspace/business_project/生产/sh.xlsx";
        staticOutputFilePath = "/Users/fury/workspace/business_project/生产/生产表格统计";
        String inputSourceImgPath = "/Users/fury/workspace/business_project/生产/p";
        staticTargetImgFilePath = "/Users/fury/workspace/business_project/生产/图库统计";


        staticSourceImgFile = new File(inputSourceImgPath);

        Path staticOutputFile = Paths.get(staticOutputFilePath);
        if (!Files.exists(staticOutputFile)) {
            try {
                Files.createDirectories(staticOutputFile);
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }

        newImgFile = getAllImages(staticSourceImgFile, new ArrayList<>()).toArray(new File[0]);

        readExcel(inputSourceExcelPath);



//        scanner.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    public static void readExcel(String inputSourceExcelPath) {
        System.out.println("====开始读取原表");
        List<List<Object>> allList = new ArrayList<>();
        try (FileInputStream inputStream = new FileInputStream(Paths.get(inputSourceExcelPath).toFile());
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            // 获取第一个工作表
            Sheet sheet = workbook.getSheetAt(0);

            // 遍历行
            for (Row row : sheet) {
                List<Object> list = new ArrayList<>();
                // 遍历列
                for (Cell cell : row) {
                    // 根据单元格类型获取内容
                    switch (cell.getCellType()) {
                        case STRING:
                            list.add(cell.getStringCellValue() == null ? "" : cell.getStringCellValue());
                            break;
                        case NUMERIC:
                            Integer count = (int) cell.getNumericCellValue();
                            list.add(count == null ? 0 : count);
                            break;
                    }
                }

                allList.add(list);
            }
            allList.remove(0);
        } catch (IOException e) {
            e.printStackTrace();
        }
        if (allList.size() != 0) {
            handleAllList(allList);
        }
    }

    public static void handleAllList(List<List<Object>> allList) {
        List<List<Object>> newLists = new ArrayList<>();
        List<Object> list;
        List<Object> tempList;
        for (int i = 0; i < allList.size(); i++) {
            list = allList.get(i);

            // 包含+号的拆分
            String sku = String.valueOf(list.get(2));
            String[] skus = sku.split("\\+");
            int count = (int) list.get(3);
            for (int j = 0; j < skus.length; j++) {
                list.set(2, skus[j]);
                // 数量大于1拆分
//                if (count > 1) {
                    list.set(3, 1);
                    tempList = new ArrayList<>(list);
                    newLists.add(tempList);
                    for (int k = 1; k < count; k++) {
                        newLists.add(tempList);
                    }
//                } else {
//                    newLists.add(list);
//                }
            }
        }

        // 相同订单的分在一起
        String key;
        LinkedHashMap<String, List<List<Object>>> sameOrderIdObj = new LinkedHashMap<>();
        for (int i = 0; i < newLists.size(); i++) {
            key = (String) newLists.get(i).get(1);
            if (!sameOrderIdObj.containsKey(key)) {
                List<List<Object>> newList = new ArrayList<>();
                newList.add(newLists.get(i));
                sameOrderIdObj.put(key, newList);
            } else {
                sameOrderIdObj.get(key).add(newLists.get(i));
            }
        }

        splitMainTable(sameOrderIdObj);
    }



    // 拆分创建主表
    public static void splitMainTable(LinkedHashMap<String, List<List<Object>>> sameOrderIdObj) {
        System.out.println("====开始拆表");
        // 拆分规则，每 100 个 订单号 为一个表格
        List<LinkedHashMap<String, List<List<Object>>>> newSameOrderIdObjs = new LinkedList<>();

//        if (sameOrderIdObj.size() >= 150) {
//            newSameOrderIdObjs = splitMap(sameOrderIdObj, 100, 50);
            newSameOrderIdObjs = splitLinkedHashMap(sameOrderIdObj, 50, 30);
//        } else {
//            newSameOrderIdObjs.add(sameOrderIdObj);
//        }

        System.out.println("====拆分表格共 " + newSameOrderIdObjs.size() + " 个");
        for (int i = 0; i < newSameOrderIdObjs.size(); i++) {
            System.out.println("====开始转换表 " + (i + inputTableIndex));
            handleSingleTable(newSameOrderIdObjs.get(i), i);
        }

        createStatisticsTable();
    }

    public static List<LinkedHashMap<String, List<List<Object>>>> splitLinkedHashMap(LinkedHashMap<String, List<List<Object>>> linkedHashMap, int groupSize, int minGroupSize) {
        List<LinkedHashMap<String, List<List<Object>>>> resultList = new ArrayList<>();
        List<Map.Entry<String, List<List<Object>>>> entries = new ArrayList<>(linkedHashMap.entrySet());

        // 计算需要拆分的组数
        int numGroups = (int) Math.ceil((double) linkedHashMap.size() / groupSize);

        // 循环拆分LinkedHashMap并添加到List中
        for (int i = 0; i < numGroups; i++) {
            int start = i * groupSize;
            int end = Math.min((i + 1) * groupSize, linkedHashMap.size());

            // 创建一个新的Map用于存放拆分后的结果
            LinkedHashMap<String, List<List<Object>>> groupMap = new LinkedHashMap<>();
            for (int j = start; j < end; j++) {
                Map.Entry<String, List<List<Object>>> entry = entries.get(j);
                groupMap.put(entry.getKey(), entry.getValue());
            }

            // 如果是最后一组并且不满足最小组大小要求，则将其加入到前一组中
            if (i == numGroups - 1 && groupMap.size() < minGroupSize && resultList.size() > 0) {
                Map<String, List<List<Object>>> lastGroupMap = resultList.get(resultList.size() - 1);
                lastGroupMap.putAll(groupMap);
            } else {
                resultList.add(groupMap);
            }
        }

        return resultList;
    }


    private static Map<String, Object[]> stylesMap = new HashMap<>();
    private static final Pattern ENGLISH_PATTERN = Pattern.compile("^[a-zA-Z]+$");
    public static boolean isEnglish(String str) {
        if (str == null) {
            return false;
        }
        Matcher matcher = ENGLISH_PATTERN.matcher(str);
        return matcher.matches();
    }

    // 处理单表
    public static void handleSingleTable(Map<String, List<List<Object>>> singleTable, int tableIndex) {
        List<List<Object>> temp;
        List<List<Object>> allRows = new ArrayList<>();

        // 先合并成总行
        for (String orderId : singleTable.keySet()) {
            temp = singleTable.get(orderId);
            for (List<Object> row : temp) {
                allRows.add(row);
            }
        }

        Set<String> orderIdSet = new HashSet<>();
        List<String> orderIdList = new ArrayList<>();
        List<int[]> styleIdIndexList = new ArrayList<>();
        List<Object> row;
        JSONArray rowsArray = new JSONArray();
        for (int i = 0; i < allRows.size(); i++) {
            JSONObject rowObj = new JSONObject();
            row = allRows.get(i);

            handleSku((String) row.get(2), rowObj);
            rowObj.set("orderId", row.get(1));
            rowObj.set("count", row.get(3));
            rowObj.set("desc", row.size() < 7 ? "" : row.get(4));

            rowObj.set("url", row.size() < 7 ? row.get(5) : row.get(6));
            rowObj.set("id", "pu");
            String styleIdStr = rowObj.getStr("styleId");
            if (!"notsure".equals(styleIdStr) && isEnglish(styleIdStr.substring(0, 1)) && isEnglish(styleIdStr.substring(1, 2))) {
                JSONObject rowObj2 = new JSONObject();
                String styleId = rowObj.getStr("styleId");
                rowObj2.set("orderId", rowObj.getStr("orderId"));
                rowObj2.set("styleId", styleId + "-1");
                rowObj2.set("color", rowObj.getStr("color"));
                rowObj2.set("size", rowObj.getStr("size"));
                rowObj2.set("count", rowObj.getInt("count"));
                rowObj2.set("desc", rowObj.getStr("desc"));
                rowObj2.set("url", rowObj.getStr("url"));
                rowObj2.set("id", "pu");
                rowObj2.set("type", rowObj.getStr("type"));

                rowsArray.add(rowObj2);

                rowObj.set("styleId", styleId + "-2");

                orderIdList.add(rowObj2.getStr("orderId"));

                int[] ints = new int[2];
                if (rowsArray.size() == 0) {
                    ints[0] = 0;
                    ints[1] = 1;
                } else {
                    ints[0] = rowsArray.size();
                    ints[1] = rowsArray.size() + 1;
                }

                styleIdIndexList.add(ints);
            }

            rowsArray.add(rowObj);
            orderIdList.add(rowObj.getStr("orderId"));
            orderIdSet.add(rowObj.getStr("orderId"));

            // 统计订单衣服种类
            statisticsStyles((String) row.get(0), (String) row.get(1), rowObj, stylesMap, (tableIndex + inputTableIndex));
        }

        String tableName = tablePreName + date + "-" + (tableIndex + inputTableIndex) + "-" + orderIdSet.size() + "单" + "-" + allRows.size() + "件";

        Path subFolder = Paths.get(staticTargetImgFilePath).resolve(tableName);

        try {
            if (!Files.exists(subFolder)) {
                Files.createDirectories(subFolder);
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        for (int i = 0; i < rowsArray.size(); i++) {
            JSONObject imgRow = rowsArray.getJSONObject(i);
            imgRow.set("id", tablePreName + date + "-" + (tableIndex + inputTableIndex) + "-" + (i + 1) );

            System.out.println("====开启搜图任务：" + imgRow.getStr("id"));
            System.out.println(imgRow.get("styleId"));
//            new Thread(() -> renameAndCopyImages(staticSourceImgFile, subFolder, imgRow.getStr("styleId"), imgRow.getStr("id"))).start();
//            new Thread(() -> renameAndCopyImages(staticSourceImgFile, subFolder, imgRow, tableName)).start();
            renameAndCopyImages(subFolder, imgRow, tableName);
        }

        new Thread(() -> createMainTable(getSameIndexes(orderIdSet, orderIdList), styleIdIndexList, rowsArray, tableName)).start();
    }

    public static void statisticsStyles(String key, String ydId, JSONObject sku, Map<String, Object[]> styles, int tableIndex) {
        // 【订单号】--然后按【衣服种类】统计数量
        if (!styles.containsKey(key)) {
            styles.put(key, new Object[]{ydId, 0, 0, 0, 0, 0, 0, 0, tablePreName +  date + "-" + tableIndex});
        }

        switch (sku.getStr("type").toLowerCase()) {
            case "100%cotton":
                styles.get(key)[1] = (int) styles.get(key)[1] + 1;
                break;
            case "short":
                styles.get(key)[3] = (int) styles.get(key)[3] + 1;
                break;
            case "hoodie":
                styles.get(key)[4] = (int) styles.get(key)[4] + 1;
                break;
            case "s":
                styles.get(key)[5] = (int) styles.get(key)[5] + 1;
                break;
            case "t-shirt":
                styles.get(key)[6] = (int) styles.get(key)[6] + 1;
                break;
            case "child":
                styles.get(key)[7] = (int) styles.get(key)[7] + 1;
                break;
        }
        String styleId = sku.getStr("styleId");
        if (!"notsure".equals(styleId) && isEnglish(styleId.substring(0, 1)) && isEnglish(styleId.substring(1, 2))) {
            styles.get(key)[2] = (int) styles.get(key)[2] + 1;
        }
    }

    public static List<int[]> getSameIndexes(Set<String> orderIdSet, List<String> orderIdList) {
        // 遍历set，将每个数字的首次和最后一次出现的下标放入int[]，并检查是否不同
        List<int[]> resultList = new ArrayList<>();
        // 遍历list，同时检查元素是否在set中
        for (int i = 0; i < orderIdList.size(); i++) {
            String temp = orderIdList.get(i);
            if (orderIdSet.contains(temp)) { // 如果元素在set中（即它是唯一的）
                int[] indices = findFirstAndLastIndex(orderIdList, temp); // 查找首次和最后一次出现的下标
                if (indices[0] != indices[1]) { // 如果首次和最后一次出现的下标不同
                    resultList.add(indices); // 添加到结果列表中
                }
                orderIdSet.remove(temp); // 从set中移除已处理的元素，避免重复检查
            }
        }

        return resultList;
    }

    // 在list中查找数字number的首次和最后一次出现的下标
    private static int[] findFirstAndLastIndex(List<String> list, String temp) {
        int firstIndex = -1;
        int lastIndex = -1;
        for (int i = 0; i < list.size(); i++) {
            if (list.get(i).equals(temp)) {
                if (firstIndex == -1) {
                    firstIndex = i; // 记录首次出现的下标
                }
                lastIndex = i; // 更新最后一次出现的下标
            }
        }
        return new int[]{firstIndex, lastIndex};
    }

    public static void createMainTable(List<int[]> idxs, List<int[]> styleIdIndexList, JSONArray rowsArray, String tableName) {
        System.out.println("====开启创建生产表任务："+tableName);
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("结果");

        CellStyle style = workbook.createCellStyle();
        style.setWrapText(true); // 允许文本自动换行
        style.setAlignment(HorizontalAlignment.CENTER); // 文本居中
        style.setVerticalAlignment(VerticalAlignment.CENTER); // 文本垂直居中

        // ['运单号', '款号', '颜色', '尺码', '产品数量', '订单备注', '衣服种类','图片', '编号',]

        Row headerRow = sheet.createRow(0);
        CellStyle colorStyle = workbook.createCellStyle();
        colorStyle.setWrapText(true); // 允许文本自动换行
        colorStyle.setAlignment(HorizontalAlignment.CENTER); // 文本居中
        colorStyle.setVerticalAlignment(VerticalAlignment.CENTER); // 文本垂直居中
        XSSFColor myColor = new XSSFColor(new java.awt.Color(0,176,240), null);
        colorStyle.setFillForegroundColor(myColor); //设置填充颜色
        colorStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND); //设置填充模式为实心
        setCellStyleAndValue(headerRow, 0, colorStyle, "蓝色：用阿叔的货并且撕标\n\rสีฟ้า：ด้วยผลิตภัณฑ์ของลุงและต้องตัดฉลากออก");

        CellStyle colorStyle2 = workbook.createCellStyle();
        colorStyle2.setWrapText(true); // 允许文本自动换行
        colorStyle2.setAlignment(HorizontalAlignment.CENTER); // 文本居中
        colorStyle2.setVerticalAlignment(VerticalAlignment.CENTER); // 文本垂直居中
        XSSFColor myColor2 = new XSSFColor(new java.awt.Color(112,48,160), null);
        colorStyle2.setFillForegroundColor(myColor2); //设置填充颜色
        colorStyle2.setFillPattern(FillPatternType.SOLID_FOREGROUND); //设置填充模式为实心
        setCellStyleAndValue(headerRow, 1, colorStyle2,"紫色：短款T恤\nสีม่วง: เสื้อยืดตัวสั้น\n'ခရမ်းရောင်- တီရှပ်အတို");

        CellStyle colorStyle3 = workbook.createCellStyle();
        colorStyle3.setWrapText(true); // 允许文本自动换行
        colorStyle3.setAlignment(HorizontalAlignment.CENTER); // 文本居中
        colorStyle3.setVerticalAlignment(VerticalAlignment.CENTER); // 文本垂直居中
        XSSFColor myColor3 = new XSSFColor(new java.awt.Color(191,143,0), null);
        colorStyle3.setFillForegroundColor(myColor3); //设置填充颜色
        colorStyle3.setFillPattern(FillPatternType.SOLID_FOREGROUND); //设置填充模式为实心
        setCellStyleAndValue(headerRow, 2, colorStyle3,"黄色：需要撕标\nสีเหลือง: ต้องฉีกเครื่องหมาย\nအဝါရောင်- တံဆိပ်ကို ဖြတ်ပစ်ရန် လိုအပ်သည်။");

        CellStyle colorStyle4 = workbook.createCellStyle();
        colorStyle4.setWrapText(true); // 允许文本自动换行
        colorStyle4.setAlignment(HorizontalAlignment.CENTER); // 文本居中
        colorStyle4.setVerticalAlignment(VerticalAlignment.CENTER); // 文本垂直居中
        XSSFColor myColor4 = new XSSFColor(new java.awt.Color(146,208,80), null);
        colorStyle4.setFillForegroundColor(myColor4); //设置填充颜色
        colorStyle4.setFillPattern(FillPatternType.SOLID_FOREGROUND); //设置填充模式为实心
        setCellStyleAndValue(headerRow, 3, colorStyle4,"绿色：用成品\nสีเขียว: ด้วยผลิตภัณฑ์สำเร็จรูป\nအစိမ္းေရာင္- ကုန်ချောကိုသုံးပါ။");

        CellStyle colorStyle5 = workbook.createCellStyle();
        colorStyle5.setWrapText(true); // 允许文本自动换行
        colorStyle5.setAlignment(HorizontalAlignment.CENTER); // 文本居中
        colorStyle5.setVerticalAlignment(VerticalAlignment.CENTER); // 文本垂直居中
        XSSFColor myColor5 = new XSSFColor(new java.awt.Color(47,117,181), null);
        colorStyle5.setFillForegroundColor(myColor5); //设置填充颜色
        colorStyle5.setFillPattern(FillPatternType.SOLID_FOREGROUND); //设置填充模式为实心
        setCellStyleAndValue(headerRow, 4, colorStyle5,"蓝色：卫衣\nสีน้ำเงิน: เสื้อฮู้ดี้\nအပြာ : အင်္ကျီ ၊");

        setCellStyleAndValue(headerRow, 5, style,"");
        setCellStyleAndValue(headerRow, 6, style,"");
        setCellStyleAndValue(headerRow, 7, style,"");
        setCellStyleAndValue(headerRow, 8, style,"");
        headerRow.setHeightInPoints(86);

        Row row = sheet.createRow(1);
        setCellStyleAndValue(row, 0, style, "运单号");
        setCellStyleAndValue(row, 1, style, "款号");
        setCellStyleAndValue(row, 2, style, "颜色");
        setCellStyleAndValue(row, 3, style, "尺码");
        setCellStyleAndValue(row, 4, style, "产品数量");
        setCellStyleAndValue(row, 5, style, "订单备注");
        setCellStyleAndValue(row, 6, style, "衣服种类");
        setCellStyleAndValue(row, 7, style, "图片");
        setCellStyleAndValue(row, 8, style, "编号");
        row.setHeightInPoints(30);

        JSONObject rowObj;
        for (int i = 0; i < rowsArray.size(); i++) {
            rowObj = rowsArray.getJSONObject(i);
            row = sheet.createRow(i + 2);
            setCellStyleAndValue(row, 0, style, rowObj.getStr("orderId"));
            setCellStyleAndValue(row, 1, style, rowObj.getStr("styleId"));
            setCellStyleAndValue(row, 2, style, rowObj.getStr("color"));
            setCellStyleAndValue(row, 3, style, rowObj.getStr("size"));
            setCellStyleAndValue(row, 4, style, rowObj.getStr("count"));
            setCellStyleAndValue(row, 5, style, rowObj.getStr("desc"));
            setCellStyleAndValue(row, 6, getColorStyle(workbook, rowObj.getStr("type")), rowObj.getStr("type"));
            setCellStyleAndValue(row, 7, style, rowObj.getStr("url"));

            setImg(rowObj.getStr("url"), 7, i + 2,workbook, sheet);

            setCellStyleAndValue(row, 8, style, rowObj.getStr("id"));

            row.setHeightInPoints(70);
        }

        // 合并运单号
        for (int[] idx : idxs) {
            sheet.addMergedRegion(mergeRegion(idx[0] + 2, idx[1] + 2, 0, 0));
        }

        // 根据款号合并颜色尺码数量衣服种类
        for (int[] idx : styleIdIndexList) {
            sheet.addMergedRegion(mergeRegion(idx[0] + 1, idx[1] + 1, 2, 2));
            sheet.addMergedRegion(mergeRegion(idx[0] + 1, idx[1] + 1, 3, 3));
            sheet.addMergedRegion(mergeRegion(idx[0] + 1, idx[1] + 1, 4, 4));
            sheet.addMergedRegion(mergeRegion(idx[0] + 1, idx[1] + 1, 6, 6));
        }


        sheet.setColumnWidth(0, 24 * 256);
        sheet.setColumnWidth(1, 20 * 256);
        sheet.setColumnWidth(2, 41 * 256);
        sheet.setColumnWidth(3, 22 * 256);
        sheet.setColumnWidth(4, 13 * 256);
        sheet.setColumnWidth(5, 10 * 256);
        sheet.setColumnWidth(6, 12 * 256);
        sheet.setColumnWidth(7, 12 * 256);
        sheet.setColumnWidth(8, 12 * 256);

        try (FileOutputStream outputStream = new FileOutputStream(staticOutputFilePath + "\\" + tableName + ".xlsx")){
            workbook.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public static CellStyle getColorStyle(Workbook workbook, String type) {
        String rgb = type2ColorMap.get(type.toLowerCase());
        CellStyle colorStyle = workbook.createCellStyle();
        colorStyle.setWrapText(true); // 允许文本自动换行
        colorStyle.setAlignment(HorizontalAlignment.CENTER); // 文本居中
        colorStyle.setVerticalAlignment(VerticalAlignment.CENTER); // 文本垂直居中
        if (rgb != null) {
            String[] split = rgb.split(",");
            colorStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(Integer.parseInt(split[0]), Integer.parseInt(split[1]), Integer.parseInt(split[2])), null)); //设置填充颜色
            colorStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND); //设置填充模式为实心
        }
        return colorStyle;
    }

    public static void setCellStyleAndValue(Row row, int index, CellStyle style, String v) {
        Cell cell = row.createCell(index);
        cell.setCellStyle(style);
        cell.setCellValue(v);
    }

    public static void setImg(String urlStr, int col, int row, Workbook workbook, Sheet sheet) {
        try {
            // 从 URL 下载图片
            System.out.println("urlStr: " + urlStr);
            URL url = new URL(urlStr); // 替换为实际的图片 URL
            URLConnection connection = url.openConnection();
            InputStream inputStream = connection.getInputStream();
            byte[] bytes = IOUtils.toByteArray(inputStream);
            inputStream.close();

            // 将图片转换为 BufferedImage
            InputStream bis = new ByteArrayInputStream(bytes);
            BufferedImage bufferedImage = ImageIO.read(bis);
            bis.close();

            // 将图片添加到工作簿中
            int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);

            // 创建辅助对象以设置图片位置
            CreationHelper helper = workbook.getCreationHelper();
            Drawing drawing = sheet.createDrawingPatriarch();
            ClientAnchor anchor = helper.createClientAnchor();

            // 设置图片位置，使其位于特定单元格上
            // 注意：Excel 中的图片位置是基于行和列的偏移量，而不是直接绑定到单元格
            // 下面的代码将图片放在 A1 单元格上，你可能需要调整 dx1, dy1, dx2, dy2 的值以达到最佳效果
            anchor.setCol1(col); // 第一列
            anchor.setRow1(row); // 第一行
            anchor.setDx1(0); // x 偏移量
            anchor.setDy1(0); // y 偏移量
            anchor.setCol2(col + 1); // 下一列（图片宽度）
            anchor.setRow2(row + 1); // 下一行（图片高度）

            // 创建图片并设置其位置
            Picture pict = drawing.createPicture(anchor, pictureIdx);
            double scale = 0.5; // 例如，缩小到原始大小的50%
            pict.resize(1, 1); // 调整图片大小以适应单元格（可选）
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static CellRangeAddress mergeRegion(int firstRow, int lastRow, int firstColumn, int lastColumn) {
        return new CellRangeAddress(
                firstRow, // first row (0-based)
                lastRow, // last row  (0-based)
                firstColumn, // first column (0-based)
                lastColumn  // last column  (0-based)
        );
    }

    private static Map<String, Integer> miniSkusCount = new HashMap<>();
    private static Map<String, Integer> skuIdsCount = new HashMap<>();

    private static Map<String, String> color2StrMap = new HashMap<>();
    private static Map<String, String> type2StrMap = new HashMap<>();

    private static Map<String, String> preTypeMap = new HashMap<>();
    private static Map<String, String> type2ColorMap = new HashMap<>();
    static {
        color2StrMap.put("White","\"白色/White/สีขาว/အဖြူ\"");
        color2StrMap.put("Rose Red","\"玫红色/Rose Red/กุหลาบแดง/အနီရောင်\"");
        color2StrMap.put("Dark Grey","\"深灰色/Dark Grey/สีเทาเข้ม/မီးခိုေရာင္\"");
        color2StrMap.put("Pink","\"粉红色/Pink/สีชมพู/ပန္ေရာင္\"");
        color2StrMap.put("Black","\"黑色/Black/สีดำ/အမည်း\"");
        color2StrMap.put("black","\"黑色/Black/สีดำ/အမည်း\"");
        color2StrMap.put("Army Green","\"军绿色/Army Green/อาร์มี่กรีน/အစိမ္းေရာင္\"");
        color2StrMap.put("Brown","\"棕色/Brown/สีน้ำตาล/ငပိအေရာင္\"");
        color2StrMap.put("Apricot","\"杏色/Apricot/แอปริคอท/Apricot\"");
        color2StrMap.put("Red","\"红色/Red/สีแดง/အနီေရာင်\"");
        color2StrMap.put("Wine Red","\"酒红色/Wine Red/สีแดงเลือดหมู/ဘာဂန်ဒီ။\"");
        color2StrMap.put("Navy Blue","\"藏青色/Navy Blue/น้ำเงิน/အပြာရင့်\"");
        color2StrMap.put("Light Blue","\"浅蓝色/Light Blue/สีฟ้า/မိုးပြာရောင်\"");
        color2StrMap.put("Purple","\"紫色/Purple/สีม่วง/ဗေဒါရောင်\"");

        type2StrMap.put("100%cotton", "T-shirt");
        type2StrMap.put("short", "short");
        type2StrMap.put("hoodie", "Hoodie");
        type2StrMap.put("s", "随机成品");
        type2StrMap.put("t", "聚酯纤维");
        type2StrMap.put("child", "Child");

        preTypeMap.put("short", "100%cotton");
        preTypeMap.put("s", "随机成品");

        type2ColorMap.put("short", "112,48,160");
        type2ColorMap.put("s", "146,208,80");
        type2ColorMap.put("hoodie", "47,117,181");
    }
    public static JSONObject handleSku(String sku, JSONObject rowObj) {
        // 3个-中间加 100%Cotton
        String[] tempArray = sku.substring(0, sku.indexOf('(')).split("-");
        String kg = sku.substring(sku.indexOf('('), sku.length());
        if (tempArray.length == 3) {
            tempArray = (tempArray[0] + "-100%Cotton-" + tempArray[1] + "-" + tempArray[2]).split("-");
        }

        // AA7879-Apricot-M
        // 款号
        rowObj.set("styleId", tempArray[0]);
        // 衣服种类
        rowObj.set("type", tempArray[1]);
        // 颜色
        rowObj.set("color", color2StrMap.get(tempArray[2]));
        // 尺码
        rowObj.set("size",tempArray[3]);

        // 统计衣服
        statisticsCount(tempArray[1] + "-" + tempArray[2] + "-" + tempArray[3] + kg, miniSkusCount);
        // 统计款号
        statisticsCount(tempArray[0], skuIdsCount);

        return rowObj;
    }

    public static void statisticsCount(String key, Map<String, Integer> map) {
        if (!map.containsKey(key)) {
            map.put(key,1);
        } else {
            int count = map.get(key);
            count += 1;
            map.put(key, count);
        }
    }

    public static String preType(String key) {
        String[] split = key.split("-");
        String temp = split[0].toLowerCase();
        if ("short".equals(temp) || "s".equals(temp)) {
            key = preTypeMap.get(temp) + "-" + key;
        }
        return key;
    }

    public static void createStatisticsTable() {
        System.out.println("====开始创建统计表");
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("衣服种类+颜色+尺码");
        Row row = sheet.createRow(0);
        row.createCell(0).setCellValue("类型");
        row.createCell(1).setCellValue("尺码");
        row.createCell(2).setCellValue("数量");
        int sheetRowCount = 1;
        for (String key : miniSkusCount.keySet()) {
            Row row1 = sheet.createRow(sheetRowCount);
            row1.createCell(0).setCellValue(type2StrMap.get(key.split("-")[0].toLowerCase()));
            row1.createCell(1).setCellValue(preType(key));
            row1.createCell(2).setCellValue(miniSkusCount.get(key));
            sheetRowCount += 1;
        }

        Sheet sheet2 = workbook.createSheet("订单号+衣服种类");
        Row row2 = sheet2.createRow(0);
        row2.createCell(0).setCellValue("订单号");
        row2.createCell(1).setCellValue("运单号");
        row2.createCell(2).setCellValue("T恤数量");
        row2.createCell(3).setCellValue("双面数量");
        row2.createCell(4).setCellValue("短款T恤数量");
        row2.createCell(5).setCellValue("卫衣数量");
        row2.createCell(6).setCellValue("成品数量");
        row2.createCell(7).setCellValue("聚酯纤维数量");
        row2.createCell(8).setCellValue("童装数量");
        row2.createCell(9).setCellValue("所在表格");
        sheetRowCount = 1;
        for (String key : stylesMap.keySet()) {
            Row row21 = sheet2.createRow(sheetRowCount);
            row21.createCell(0).setCellValue(key);
            row21.createCell(1).setCellValue(String.valueOf(stylesMap.get(key)[0]));
            row21.createCell(2).setCellValue(String.valueOf(stylesMap.get(key)[1]));
            row21.createCell(3).setCellValue(String.valueOf(stylesMap.get(key)[2]));
            row21.createCell(4).setCellValue(String.valueOf(stylesMap.get(key)[3]));
            row21.createCell(5).setCellValue(String.valueOf(stylesMap.get(key)[4]));
            row21.createCell(6).setCellValue(String.valueOf(stylesMap.get(key)[5]));
            row21.createCell(7).setCellValue(String.valueOf(stylesMap.get(key)[6]));
            row21.createCell(8).setCellValue(String.valueOf(stylesMap.get(key)[7]));
            row21.createCell(9).setCellValue(String.valueOf(stylesMap.get(key)[8]));
            sheetRowCount += 1;
        }

        int doubleCount = 0;
        Sheet sheet3 = workbook.createSheet("款号");
        Row row3 = sheet3.createRow(0);
        row3.createCell(0).setCellValue("款号");
        row3.createCell(1).setCellValue("数量");
        sheetRowCount = 1;
        for (String key : skuIdsCount.keySet()) {
            Row row31 = sheet3.createRow(sheetRowCount);
            row31.createCell(0).setCellValue(key);
            row31.createCell(1).setCellValue(skuIdsCount.get(key));
            sheetRowCount += 1;
            if (!"notsure".equals(key) && isEnglish(key.substring(0, 1)) && isEnglish(key.substring(1, 2))) {
                doubleCount += skuIdsCount.get(key);
            }
        }


        Sheet sheet4 = workbook.createSheet("双面总数");
        Row headerRow4 = sheet4.createRow(0);
        headerRow4.createCell(0).setCellValue(doubleCount);

        Sheet sheet5 = workbook.createSheet("丢失图片");
        Row row5 = sheet5.createRow(0);
        row5.createCell(0).setCellValue("款号");
        row5.createCell(1).setCellValue("序号");
        row5.createCell(2).setCellValue("所在表格");
        JSONObject object;
        sheetRowCount = 1;
        for (int i = 0; i < missImgArray.size(); i++) {
            object = missImgArray.getJSONObject(i);
            Row row51 = sheet5.createRow(sheetRowCount);
            row51.createCell(0).setCellValue(object.getStr("styleId"));
            row51.createCell(1).setCellValue(object.getStr("id"));
            row51.createCell(2).setCellValue(object.getStr("tableName"));
            sheetRowCount += 1;
        }


        try (FileOutputStream outputStream = new FileOutputStream(Paths.get(staticOutputFilePath + "\\统计.xlsx").toFile())){
            workbook.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private static JSONArray missImgArray = new JSONArray();

    private static void renameAndCopyImages(Path targetImgPath, JSONObject imgRow, String tableName) {
        JSONObject object = searchImg(imgRow);
        Boolean success = object.getBool("success");
        String oldName = imgRow.getStr("styleId");
        String newName = imgRow.getStr("id");

        if (success) {
            missImgArray.add(new JSONObject().set("styleId", oldName).set("id", newName).set("tableName", tableName));
        } else {
            File file = new File(object.getStr("file"));
            Path destinationPath = new File(targetImgPath.toFile(), imgRow.getStr("id") + file.getName().substring(file.getName().lastIndexOf("."))).toPath();
            try {
                Files.copy(file.toPath(), destinationPath, StandardCopyOption.REPLACE_EXISTING);
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }
    }

    public static List<File> getAllImages(File sourceImgFolder, List<File> imageFiles) {
        File[] oldFiles = sourceImgFolder.listFiles();
        // 遍历旧数组，检查每个文件是否是图片
        for (File file : oldFiles) {
            if (file.isDirectory()) {
                getAllImages(file, imageFiles);
            } else if (isImageFile(file)) {
                imageFiles.add(file); // 如果是图片，添加到集合中
            }
        }
        return imageFiles;
    }

    public static JSONObject searchImg(JSONObject imgRow) {
        JSONObject object = new JSONObject();
        object.set("success", true);
        for (File file : newImgFile) {
            if (isImageFile(file)) {
                String oldName = imgRow.getStr("styleId");
                try {
                    if (oldName.equals(file.getName().substring(0, file.getName().lastIndexOf(".")))) {
                        object.set("success", false).set("file", file);
                        return object;
                    }
                } catch (Exception e) {
                    throw new RuntimeException(e);
                }
            }
        }
        return object;
    }

    private static boolean isImageFile(File file) {
        String fileName = file.getName().toLowerCase();
        return fileName.endsWith(".jpg") || fileName.endsWith(".jpeg") || fileName.endsWith(".png");
    }
}