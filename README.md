# POIDemo
Java使用POI读取试题的docx文件

## 一、准备的docx文件模板样式
![demo注意事项](./src/main/resources/demo.png)
文件中的文字挨着左边写，要是有缩进/空格之类的字符在，poi读取时不好匹配到想要的数据。
## 二、引入poi依赖
```xml
    <dependencies>
        <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi-ooxml</artifactId>
            <version>4.1.2</version>
        </dependency>
        <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi-scratchpad</artifactId>
            <version>4.1.2</version>
        </dependency>
    </dependencies>
```
## 三、编写代码
```java
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

        File file = new File("C:\\Users\\Administrator\\Desktop\\poi练习文档.docx");
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
            if (matcher1.matches()) {// 匹配到了试卷标题名
                System.out.println("试题名："+trim);               
            } else if (matcher.matches()) {// 匹配到了问题描述 
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
            } 
        }
        System.out.println(list.size()+"----读取内容----"+list.toString());
    }

}
```
## 自己做的实体类
```java
/**
 * 实体类
 * Create by Zoe on 2020/03/31 22:41
 */
public class SingleChoice {
    private String id;
    private String question;
    private String optionA;
    private String optionB;
    private String optionC;
    private String optionD;
    private String answer;
    private String analysis;
    private int type;

    @Override
    public String toString() {
        return "SingleChoice{" +
                "id='" + id + '\'' +
                ", question='" + question + '\'' +
                ", optionA='" + optionA + '\'' +
                ", optionB='" + optionB + '\'' +
                ", optionC='" + optionC + '\'' +
                ", optionD='" + optionD + '\'' +
                ", answer='" + answer + '\'' +
                ", analysis='" + analysis + '\'' +
                ", type=" + type +
                '}';
    }

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getQuestion() {
        return question;
    }

    public void setQuestion(String question) {
        this.question = question;
    }

    public String getOptionA() {
        return optionA;
    }

    public void setOptionA(String optionA) {
        this.optionA = optionA;
    }

    public String getOptionB() {
        return optionB;
    }

    public void setOptionB(String optionB) {
        this.optionB = optionB;
    }

    public String getOptionC() {
        return optionC;
    }

    public void setOptionC(String optionC) {
        this.optionC = optionC;
    }

    public String getOptionD() {
        return optionD;
    }

    public void setOptionD(String optionD) {
        this.optionD = optionD;
    }

    public String getAnswer() {
        return answer;
    }

    public void setAnswer(String answer) {
        this.answer = answer;
    }

    public String getAnalysis() {
        return analysis;
    }

    public void setAnalysis(String analysis) {
        this.analysis = analysis;
    }

    public int getType() {
        return type;
    }

    public void setType(int type) {
        this.type = type;
    }
}
```
