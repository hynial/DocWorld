package me.caseuse.workpayloadsplit;

import me.util.DocUtil;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.*;
import java.nio.file.Files;
import java.text.SimpleDateFormat;
import java.util.*;

public class GmMain {

    /**
     * 有漏洞，并行三人后，一个人周一做的事情依赖于另一个人周尾才能完成。
     * @param args
     */
    public static void main(String[] args) {
        String sourceDir = "/Users/hynial/PreArchiveCorps/RSZH/金智项目/8.工作汇报/成果分享";
        File[] sourceFiles = collectFiles(new File(sourceDir), ".*\\.docx$");
        String outDir = "/Users/hynial/PreArchiveCorps/RSZH/金智项目/8.工作汇报/成果分享/out";
        File outDirFile = new File(outDir);
        if (!outDirFile.exists()) {
            outDirFile.mkdirs();
        }
        long total = daysBetween("20241112", "20241227");
        long daysZhao = daysBetween("20240927", "20241227");
        long daysXu = daysBetween("20241108", "20241227");
        System.out.println(daysZhao + "," + daysXu + "," + total);
//        System.exit(0);
        List<File> files = Arrays.asList(sourceFiles);
        Collections.sort(files, new Comparator<File>() {
            @Override
            public int compare(File o1, File o2) {
                return o1.getName().compareTo(o2.getName());
            }
        });
        for(int i = 0; i < files.size(); i++) {
            File f = files.get(i);
            // System.out.println(f.getName());

            String fileName = f.getName().substring(0, f.getName().lastIndexOf("."));

            String namePart1 = fileName.split("-")[0];
            String remainPart = fileName.split("-")[1];
            String namePart2 = remainPart.split("_")[0];
            String namePart3 = remainPart.split("_")[1];
            String namePart4 = remainPart.split("_")[2];

//            if (namePart4.compareTo("20240927") <= 0) { // 20241115 20241227
//                namePart2 = "陈兆欣";
//                namePart4 = addDays(namePart4, (int) daysZhao);
//            } else if (namePart4.compareTo("20241108") <= 0) {
//                namePart2 = "许云娇";
//                namePart4 = addDays(namePart4, (int) daysXu);
//            } else {
//            }

            if (i % 3 == 0) {
                namePart2 = "陈兆欣";
            } else if (i % 3 == 1) {
                namePart2 = "许云娇";
            }
            String actualPart4 = namePart4;
            namePart4 = addDays("20241112", (int) (i * 1.0 / files.size() * total));

            String newName = namePart1 + "-" + namePart2 + "_" + namePart3 + "_" + namePart4;
            System.out.println(newName);

            try {
                InputStream templateStream = Files.newInputStream(f.toPath());
                XWPFDocument document = new XWPFDocument(templateStream);

                DocUtil.replaceAllDigital(document, new String[] {actualPart4}, new String[]{namePart4});

                // 将替换后的内容写入新文件
                FileOutputStream fos = new FileOutputStream(outDir + "/"+ newName +".docx");
                document.write(fos);
                fos.close();
                templateStream.close();
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }
    }

    static File[] collectFiles(File dir, final String regex) {
        File[] foundFiles = dir.listFiles(new FilenameFilter() {
            @Override
            public boolean accept(File dir, String name) {
                return regex == null || name.matches(regex);
            }
        });

        return foundFiles;
    }

    static long daysBetween(String startDay, String endDay) {
        Calendar startCalendar = Calendar.getInstance();
        startCalendar.set(Integer.parseInt(startDay.substring(0, 4)), Integer.parseInt(startDay.substring(4, 6)) - 1, Integer.parseInt(startDay.substring(6)));
        Calendar endCalendar = Calendar.getInstance();
        endCalendar.set(Integer.parseInt(endDay.substring(0, 4)), Integer.parseInt(endDay.substring(4, 6)) - 1, Integer.parseInt(endDay.substring(6)));

        // 计算两个日期之间的天数差
        long daysBetween = (endCalendar.getTimeInMillis() - startCalendar.getTimeInMillis()) / (24 * 60 * 60 * 1000);
        return daysBetween;
    }

    static String addDays(String day, int s) {
        Calendar calendar = Calendar.getInstance();
        calendar.set(Integer.parseInt(day.substring(0, 4)), Integer.parseInt(day.substring(4, 6)) - 1, Integer.parseInt(day.substring(6)));
        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd");

        // 增加s天
        calendar.add(Calendar.DAY_OF_MONTH, s);
        return sdf.format(calendar.getTime());
    }
}
