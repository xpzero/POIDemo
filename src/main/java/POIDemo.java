import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * poi测试
 * Create by Zoe on 2020/03/31 20:07
 */
public class POIDemo {

    public static void main(String[] args) throws IOException {
        // 匹配题目描述
        String regEx = "^\\d(.*?)$";
        // 匹配标题
        String regEx1 = "^\\d(.*?)答案与解读$";
        Pattern pattern = Pattern.compile(regEx);
        Pattern pattern2 = Pattern.compile(regEx1);

        File file = new File("src/main/resources/2012年全国硕士研究生统一考试政治理论试题及其答案详解.docx");
        FileInputStream stream = new FileInputStream(file);

        List<SingleChoice> list = new ArrayList<>();
        SingleChoice singleChoice = new SingleChoice();

        XWPFDocument document = new XWPFDocument(stream);
        // 读取文件中的段落，选项的一行表示一段
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            // 获取段落文字
            String trim = paragraph.getParagraphText();
            // 正则匹配
            Matcher matcher = pattern.matcher(trim);
            Matcher matcher1 = pattern2.matcher(trim);
            if (matcher.matches()) {// 匹配到了问题描述
                singleChoice.setQuestion(trim);
            } else if (trim.startsWith("A")) {
                singleChoice.setOptionA(trim);
            } else if (trim.startsWith("B")) {
                singleChoice.setOptionB(trim);
            } else if (trim.startsWith("C")) {
                singleChoice.setOptionC(trim);
            } else if (trim.startsWith("D")) {
                singleChoice.setOptionD(trim);
            } else if (trim.startsWith("答案")) {
                singleChoice.setAnswer(trim.replace("答案：", ""));
            } else if (trim.startsWith("解析")) {
                // 1. 匹配到了"解析"，表明这一题结束。
                singleChoice.setAnalysis(trim.replace("解析：", ""));
                // 2. 将这一题的数据添加到集合
                list.add(singleChoice);
                // 3. 将实体类中的数据清空,在这个for循环中载入新数据
                singleChoice = new SingleChoice();
            } else if (matcher1.matches()) {// 匹配到了试卷标题名
                System.out.println("试题名："+trim);
            }
        }
        System.out.println(list.size()+"----读取内容----"+list.toString());
    }

}
