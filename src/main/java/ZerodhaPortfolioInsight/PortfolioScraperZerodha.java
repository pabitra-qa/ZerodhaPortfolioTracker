package ZerodhaPortfolioInsight;

import com.google.gson.Gson;
import com.google.gson.JsonArray;
import com.google.gson.JsonObject;
import com.microsoft.playwright.*;
import com.microsoft.playwright.options.FormData;
import com.microsoft.playwright.options.RequestOptions;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.*;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import javax.swing.*;
import java.io.*;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.temporal.ChronoUnit;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@SuppressWarnings("ALL")
public class PortfolioScraperZerodha {
    private static final String ANSI_RESET = "\u001B[0m";
    private static final String RUPEE = "Rs. ";//"\u20B9";
    private static final String BOLD = "\033[0;1m";
    private static final String stars = "***********************************************************************************************************************************";
    private static final String hypens = "-----------------------------------------------------------------------------------------------------------------------------------";
    static PrintWriter outputfile;
    static File tempFile;
    static String tempOutputFilePath;
    static String ConsoleOutput;
    private static ArrayList<ArrayList<Object>> holdings;
    private static String mutualFundFlag;
    private static boolean sameAsPrevHoldings;
    public static JFrame frame;
    public static JTextArea textArea;
    public static XSSFCellStyle textLeftBold;
    public static XSSFCellStyle textRightBold;
    public static XSSFCellStyle textLeftNormal;
    public static XSSFCellStyle currencyRight;
    public static XSSFCellStyle intTypeNumberRight;
    public static XSSFCellStyle percentileRight;
    public static XSSFCellStyle bottomRow;
    public static String Username;
    public static String Password;
    public static String Pin;
    static boolean fl = true;
    public static XSSFCellStyle dummy;
    public static void mainMethod(String userName, char[] password, char[] pin, String coinCheck) throws IOException, InterruptedException {
        mutualFundFlag = coinCheck;
        outputfile = new PrintWriter(createTempFile());
        LocalDateTime st = LocalDateTime.now();
        Username = userName;
        Password = String.valueOf(password);
        Pin = String.valueOf(pin);
        fetchHoldings(userName, String.valueOf(password), String.valueOf(pin));
        if (fl) {
            printHoldings();
            printTop5Holdings();
            printTop5Gainers();
            printTop5Losers();
            LocalDateTime et = LocalDateTime.now();
            System.out.println("\nCompleted in " + st.until(et, ChronoUnit.SECONDS) + " seconds...");
            outputfile.close();
            ConsoleOutput = getConsoleOutputAsString(tempOutputFilePath);
            textArea.append(ConsoleOutput);
            deleteTempFile(tempFile);
            if (sameAsPrevHoldings) {
                System.out.println("\nNo updates...\n");
                textArea.append("\nNo updates...\n");
            } else {
                System.out.println("\nUpdates Found... You should write it to Excel...\n");
                textArea.append("\nUpdates Found... You should write it to Excel...\n");
            }
        }
    }

    public static void initalizeCellStyles(XSSFWorkbook workbook) {
        XSSFColor black = new XSSFColor(new java.awt.Color(255, 255, 255), null);
        XSSFColor white = new XSSFColor(new java.awt.Color(255, 255, 255), null);
        XSSFColor green = new XSSFColor(new java.awt.Color(73, 141, 33), null);

        XSSFDataFormat df = (XSSFDataFormat) workbook.createDataFormat();

        XSSFFont BoldFont = workbook.createFont();
        BoldFont.setFontName("Bahnschrift Light");
        BoldFont.setFontHeight(10);
        BoldFont.setBold(true);

        XSSFFont NormalFont = workbook.createFont();
        NormalFont.setFontName("Bahnschrift Light");
        NormalFont.setFontHeight(10);
        NormalFont.setBold(false);

        dummy = workbook.createCellStyle();
        dummy.setFillBackgroundColor(white);
        dummy.setFillForegroundColor(black);
        dummy.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        dummy.setVerticalAlignment(VerticalAlignment.CENTER);

        textLeftBold = workbook.createCellStyle();
        textLeftBold.setFillBackgroundColor(white);
        textLeftBold.setFillForegroundColor(black);
        textLeftBold.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        textLeftBold.setVerticalAlignment(VerticalAlignment.CENTER);

        textLeftBold.setAlignment(HorizontalAlignment.LEFT);
        textLeftBold.setFont(BoldFont);
        textLeftBold.setBorderBottom(BorderStyle.THIN);
        textLeftBold.setBottomBorderColor(green);

        textRightBold = workbook.createCellStyle();
        textRightBold.setFillBackgroundColor(white);
        textRightBold.setFillForegroundColor(black);
        textRightBold.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        textRightBold.setVerticalAlignment(VerticalAlignment.CENTER);

        textRightBold.setAlignment(HorizontalAlignment.RIGHT);
        textRightBold.setFont(BoldFont);
        textRightBold.setBorderBottom(BorderStyle.THIN);
        textRightBold.setBottomBorderColor(green);

        textLeftNormal = workbook.createCellStyle();
        textLeftNormal.setFillBackgroundColor(white);
        textLeftNormal.setFillForegroundColor(black);
        textLeftNormal.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        textLeftNormal.setVerticalAlignment(VerticalAlignment.CENTER);

        textLeftNormal.setAlignment(HorizontalAlignment.LEFT);
        textLeftNormal.setFont(NormalFont);

        currencyRight = workbook.createCellStyle();
        currencyRight.setFillBackgroundColor(white);
        currencyRight.setFillForegroundColor(black);
        currencyRight.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        currencyRight.setVerticalAlignment(VerticalAlignment.CENTER);

        currencyRight.setAlignment(HorizontalAlignment.RIGHT);
        currencyRight.setFont(NormalFont);
        currencyRight.setDataFormat(df.getFormat("₹ #,##0.00"));

        intTypeNumberRight = workbook.createCellStyle();
        intTypeNumberRight.setFillBackgroundColor(white);
        intTypeNumberRight.setFillForegroundColor(black);
        intTypeNumberRight.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        intTypeNumberRight.setVerticalAlignment(VerticalAlignment.CENTER);

        intTypeNumberRight.setAlignment(HorizontalAlignment.RIGHT);
        intTypeNumberRight.setFont(NormalFont);
        intTypeNumberRight.setDataFormat(df.getFormat("#,##0"));

        percentileRight = workbook.createCellStyle();
        percentileRight.setFillBackgroundColor(white);
        percentileRight.setFillForegroundColor(black);
        percentileRight.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        percentileRight.setVerticalAlignment(VerticalAlignment.CENTER);

        percentileRight.setAlignment(HorizontalAlignment.RIGHT);
        percentileRight.setFont(NormalFont);
        percentileRight.setDataFormat(df.getFormat("0.00%"));

        bottomRow = workbook.createCellStyle();
        bottomRow.setFillBackgroundColor(white);
        bottomRow.setFillForegroundColor(black);
        bottomRow.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        bottomRow.setVerticalAlignment(VerticalAlignment.CENTER);

        bottomRow.setBorderBottom(BorderStyle.THIN);
        bottomRow.setBottomBorderColor(green);
    }

    public static void setupDashBoardPage(XSSFWorkbook wb, XSSFSheet sheet) {
        initalizeCellStyles(wb);

        if (holdings.size() + 3 < 30) {
            for (int i = 0; i <= 30; i++) {
                sheet.createRow(i).setHeight((short) 315);
                for (int j = 0; j < 18; j++) {
                    sheet.getRow(i).createCell(j);
                    sheet.getRow(i).getCell(j).setCellStyle(dummy);
                }
            }
        } else {
            for (int i = 0; i <= holdings.size() + 3; i++) {
                sheet.createRow(i).setHeight((short) 315);
                for (int j = 0; j < 18; j++) {
                    sheet.getRow(i).createCell(j);
                    sheet.getRow(i).getCell(j).setCellStyle(dummy);
                }
            }
        }

        for (int i = 1; i <= 10; i++) {
            if (i >= 1 && i <= 2) {
                sheet.getRow(1).getCell(i).setCellStyle(textLeftBold);
                sheet.getRow(2).getCell(i).setCellStyle(textLeftBold);
            } else {
                sheet.getRow(1).getCell(i).setCellStyle(textRightBold);
                sheet.getRow(2).getCell(i).setCellStyle(textRightBold);
            }
            sheet.addMergedRegion(new CellRangeAddress(1, 2, i, i));
        }
        sheet.getRow(1).getCell(1).setCellValue("Ticker");
        sheet.getRow(1).getCell(2).setCellValue("Type");
        sheet.getRow(1).getCell(3).setCellValue("Quantity");
        sheet.getRow(1).getCell(4).setCellValue("Average Price");
        sheet.getRow(1).getCell(5).setCellValue("Invested Value");
        sheet.getRow(1).getCell(6).setCellValue("Current Price");
        sheet.getRow(1).getCell(7).setCellValue("Current Value");
        sheet.getRow(1).getCell(8).setCellValue("P&L");
        sheet.getRow(1).getCell(9).setCellValue("% P&L");
        sheet.getRow(1).getCell(10).setCellValue("Weight");

        for (int i = 0; i < 24; i++) {
            sheet.getRow(i).getCell(12).setCellStyle(textLeftNormal);
            sheet.getRow(i).getCell(15).setCellStyle(textLeftNormal);
        }
        sheet.getRow(4).getCell(13).setCellStyle(currencyRight);
        sheet.getRow(5).getCell(13).setCellStyle(currencyRight);
        sheet.getRow(6).getCell(13).setCellStyle(currencyRight);
        sheet.getRow(7).getCell(13).setCellStyle(currencyRight);

        sheet.getRow(4).getCell(16).setCellStyle(currencyRight);
        sheet.getRow(5).getCell(16).setCellStyle(currencyRight);
        sheet.getRow(6).getCell(16).setCellStyle(currencyRight);
        sheet.getRow(7).getCell(16).setCellStyle(percentileRight);

        sheet.getRow(12).getCell(13).setCellStyle(percentileRight);
        sheet.getRow(13).getCell(13).setCellStyle(percentileRight);
        sheet.getRow(14).getCell(13).setCellStyle(percentileRight);
        sheet.getRow(15).getCell(13).setCellStyle(percentileRight);

        sheet.getRow(12).getCell(16).setCellStyle(currencyRight);
        sheet.getRow(13).getCell(16).setCellStyle(currencyRight);
        sheet.getRow(14).getCell(16).setCellStyle(currencyRight);
        sheet.getRow(15).getCell(16).setCellStyle(percentileRight);

        sheet.getRow(20).getCell(13).setCellStyle(percentileRight);
        sheet.getRow(21).getCell(13).setCellStyle(percentileRight);
        sheet.getRow(22).getCell(13).setCellStyle(percentileRight);
        sheet.getRow(23).getCell(13).setCellStyle(percentileRight);

        sheet.getRow(20).getCell(16).setCellStyle(currencyRight);
        sheet.getRow(21).getCell(16).setCellStyle(currencyRight);
        sheet.getRow(22).getCell(16).setCellStyle(currencyRight);
        sheet.getRow(23).getCell(16).setCellStyle(percentileRight);

        for (int row = 3; row < holdings.size() + 3; row++) {
            sheet.getRow(row).getCell(1).setCellStyle(textLeftNormal);
            sheet.getRow(row).getCell(2).setCellStyle(textLeftNormal);
            sheet.getRow(row).getCell(3).setCellStyle(intTypeNumberRight);
            sheet.getRow(row).getCell(4).setCellStyle(currencyRight);
            sheet.getRow(row).getCell(5).setCellStyle(currencyRight);
            sheet.getRow(row).getCell(6).setCellStyle(currencyRight);
            sheet.getRow(row).getCell(7).setCellStyle(currencyRight);
            sheet.getRow(row).getCell(8).setCellStyle(currencyRight);
            sheet.getRow(row).getCell(9).setCellStyle(percentileRight);
            sheet.getRow(row).getCell(10).setCellStyle(percentileRight);
        }

        sheet.addMergedRegion(new CellRangeAddress(1, 2, 12, 13));
        sheet.addMergedRegion(new CellRangeAddress(9, 10, 12, 13));
        sheet.addMergedRegion(new CellRangeAddress(17, 18, 12, 13));
        sheet.addMergedRegion(new CellRangeAddress(1, 2, 15, 16));
        sheet.addMergedRegion(new CellRangeAddress(9, 10, 15, 16));
        sheet.addMergedRegion(new CellRangeAddress(17, 18, 15, 16));

        sheet.getRow(1).getCell(12).setCellStyle(textLeftBold);
        sheet.getRow(2).getCell(12).setCellStyle(textLeftBold);
        sheet.getRow(1).getCell(13).setCellStyle(textLeftBold);
        sheet.getRow(2).getCell(13).setCellStyle(textLeftBold);

        sheet.getRow(9).getCell(12).setCellStyle(textLeftBold);
        sheet.getRow(10).getCell(12).setCellStyle(textLeftBold);
        sheet.getRow(9).getCell(13).setCellStyle(textLeftBold);
        sheet.getRow(10).getCell(13).setCellStyle(textLeftBold);

        sheet.getRow(17).getCell(12).setCellStyle(textLeftBold);
        sheet.getRow(18).getCell(12).setCellStyle(textLeftBold);
        sheet.getRow(17).getCell(13).setCellStyle(textLeftBold);
        sheet.getRow(18).getCell(13).setCellStyle(textLeftBold);

        sheet.getRow(1).getCell(15).setCellStyle(textLeftBold);
        sheet.getRow(2).getCell(15).setCellStyle(textLeftBold);
        sheet.getRow(1).getCell(16).setCellStyle(textLeftBold);
        sheet.getRow(2).getCell(16).setCellStyle(textLeftBold);


        sheet.getRow(9).getCell(15).setCellStyle(textLeftBold);
        sheet.getRow(10).getCell(15).setCellStyle(textLeftBold);
        sheet.getRow(9).getCell(16).setCellStyle(textLeftBold);
        sheet.getRow(10).getCell(16).setCellStyle(textLeftBold);

        sheet.getRow(17).getCell(15).setCellStyle(textLeftBold);
        sheet.getRow(18).getCell(15).setCellStyle(textLeftBold);
        sheet.getRow(17).getCell(16).setCellStyle(textLeftBold);
        sheet.getRow(18).getCell(16).setCellStyle(textLeftBold);


        sheet.getRow(1).getCell(12).setCellValue("Top 4 Holdings");
        sheet.getRow(9).getCell(12).setCellValue("Top 4 Gainers");
        sheet.getRow(17).getCell(12).setCellValue("Top 4 Losers");
        sheet.getRow(1).getCell(15).setCellValue("Stock Portfolio Stats");
        sheet.getRow(9).getCell(15).setCellValue("Mutual Fund Portfolio Stats");
        sheet.getRow(17).getCell(15).setCellValue("Overall Portfolio Stats");

        for (int i = 1; i <= 16; i++) {
            sheet.getRow(holdings.size() + 3).getCell(i).setCellStyle(bottomRow);
        }
        sheet.setColumnWidth(0, 3 * 277);
        sheet.setColumnWidth(1, 15 * 277);
        sheet.setColumnWidth(2, 10 * 277);
        sheet.setColumnWidth(3, 8 * 277);
        sheet.setColumnWidth(4, 13 * 277);
        sheet.setColumnWidth(5, 14 * 277);
        sheet.setColumnWidth(6, 13 * 277);
        sheet.setColumnWidth(7, 13 * 277);
        sheet.setColumnWidth(8, 11 * 277);
        sheet.setColumnWidth(9, 9 * 277);
        sheet.setColumnWidth(10, 9 * 277);
        sheet.setColumnWidth(11, 3 * 277);
        sheet.setColumnWidth(12, 15 * 277);
        sheet.setColumnWidth(13, 15 * 277);
        sheet.setColumnWidth(14, 9 * 277);
        sheet.setColumnWidth(15, 15 * 277);
        sheet.setColumnWidth(16, 15 * 277);
    }

    private static final Comparator<ArrayList<Object>> sortByInvestedValue = new Comparator<ArrayList<Object>>() {
        public int compare(ArrayList<Object> pList1, ArrayList<Object> pList2) {
            double item1 = Double.parseDouble(pList1.get(4).toString());
            double item2 = Double.parseDouble(pList2.get(4).toString());
            if (pList1.get(1).equals("MF"))
                return 1;
            return Double.compare(item2, item1);
        }
    };

    private static final Comparator<ArrayList<Object>> sortByGain = new Comparator<ArrayList<Object>>() {
        public int compare(ArrayList<Object> pList1, ArrayList<Object> pList2) {
            double item1 = Double.parseDouble(pList1.get(8).toString());
            double item2 = Double.parseDouble(pList2.get(8).toString());
            if (pList1.get(1).equals("MF"))
                return 1;
            return Double.compare(item2, item1);
        }
    };

    private static final Comparator<ArrayList<Object>> sortByLoss = new Comparator<ArrayList<Object>>() {
        public int compare(ArrayList<Object> pList1, ArrayList<Object> pList2) {
            double item1 = Double.parseDouble(pList1.get(8).toString());
            double item2 = Double.parseDouble(pList2.get(8).toString());
            if (pList1.get(1).equals("MF"))
                return 1;
            return Double.compare(item1, item2);
        }
    };

    public static void sendEmail(String email, char[] password) throws IOException {
        //File htmlTemplateFile = new File("lib/template.html");
        String htmlString = "<!DOCTYPE html PUBLIC \"-//W3C//DTD HTML 4.01 Transitional//EN\" \n" +
                "\"http://www.w3.org/TR/html4/loose.dtd\">\n" +
                "<html>\n" +
                "<head>\n" +
                "<meta http-equiv=\"Content-Type\" content=\"text/html; charset=UTF-8\">\n" +
                "<title>$title</title>\n" +
                "</head>\n" +
                "<body>$body\n" +
                "</body>\n" +
                "</html>";
        String title = "Portfolio Stats";
        String body = "<pre>" + ConsoleOutput + "</pre>";
        htmlString = htmlString.replace("$title", title);
        htmlString = htmlString.replace("$body", body);
        File newHtmlFile = new File("Portfolio_Details.html");
        FileUtils.writeStringToFile(newHtmlFile, htmlString);

        String to = email;
        String from = email;
        String host = "smtp.gmail.com";

        Properties properties = System.getProperties();

        properties.put("mail.smtp.host", host);
        properties.put("mail.smtp.port", "465");
        properties.put("mail.smtp.ssl.enable", "true");
        properties.put("mail.smtp.auth", "true");

        Session session = Session.getInstance(properties, new Authenticator() {
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(email, String.valueOf(password));
            }
        });

        try {
            MimeMessage message = new MimeMessage(session);
            message.setFrom(new InternetAddress(from));
            message.addRecipient(Message.RecipientType.TO, new InternetAddress(to));

            message.setSubject("Portfolio Update!!! | "
                    + new SimpleDateFormat("yyyy-MM-dd HH:mm").format(new Timestamp(new Date().getTime())));

            BodyPart messageBodyPart1 = new MimeBodyPart();
            messageBodyPart1.setText("Hi Pabitra,\n\nJust now a build was triggered. " +
                    "PFA the HTML file for latest portfolio update.\n\n\nThank you!\nPabitra");

            BodyPart messageBodyPart2 = new MimeBodyPart();
            String filename = "Portfolio_Details.html";
            DataSource source = new FileDataSource(filename);
            messageBodyPart2.setDataHandler(new DataHandler(source));
            messageBodyPart2.setFileName(filename);

            // creating MultiPart object
            Multipart multipartObject = new MimeMultipart();
            multipartObject.addBodyPart(messageBodyPart1);
            multipartObject.addBodyPart(messageBodyPart2);

            message.setContent(multipartObject);
            Transport.send(message);
            textArea.append("\nEmail Sent... Please check your inbox...");
        } catch (MessagingException mex) {
            mex.printStackTrace();
            textArea.append("\nCouldn't Send Email... Please check mail id and app password again...");
        }
    }

    private static File createTempFile() throws IOException {
        String path = System.getProperty("user.dir");
        String prefix = "consoleOutput";
        String suffix = ".txt";
        File file = File.createTempFile(prefix, suffix, new File(path));
        tempOutputFilePath = file.getAbsolutePath();
        tempFile = file;
        return file;
    }

    private static void deleteTempFile(File file) {
        file.delete();
    }

    private static String getConsoleOutputAsString(String file) {
        StringBuilder builder = new StringBuilder();

        try (BufferedReader buffer = new BufferedReader(
                new FileReader(file))) {
            String str;
            while ((str = buffer.readLine()) != null) {
                builder.append(str).append("\n");
            }
        } catch (IOException e) {

            e.printStackTrace();
        }

        return builder.toString();
    }

    private static void addMFHoldingsToHoldings() {
        Playwright playwright1 = Playwright.create();
        Browser browser = playwright1.chromium().launch();
        BrowserContext bx = browser.newContext(new Browser.NewContextOptions().setViewportSize(1920, 1080));
        Page page = bx.newPage();
        page.navigate("https://coin.zerodha.com/");
        page.locator("span:has-text('Login')").first().click();
        page.waitForLoadState();

        page.locator("input#userid").fill(Username);
        page.locator("input#password").fill(Password);
        page.locator("button:has-text('Login')").click();
        page.waitForLoadState();
        page.locator("input#pin").fill(Pin);
        page.locator("button:has-text('Continue')").click();
        page.waitForLoadState();
        page.locator("//a[contains(text(),'Holdings')]").click();
        page.waitForLoadState();
        page.locator("//div[@class='row']").waitFor();
        int mfholdingSize = page.locator("//div[@class='row']").count();
        for (int i = 0; i < mfholdingSize; i++) {
            ArrayList<Object> mf = new ArrayList<>();
            String fundName = page.locator("div.row div.fund-name").textContent().trim();
            page.locator("//div[@class='row']").nth(i).click();

            double invValue = Double.parseDouble(page.locator("//div[contains(text(),'Invested')]//following-sibling::div").nth(i).textContent().replaceAll("\"", "").replaceAll(",", ""));
            double currValue = Double.parseDouble(page.locator("//div[contains(text(),'Current')]//following-sibling::div").nth(i).textContent().replaceAll("\"", "").replaceAll(",", ""));
            int qty = Integer.parseInt(page.locator("//div[contains(text(),'Units')]//following-sibling::div").nth(i).textContent().replaceAll("\"", "").replaceAll(",", ""));
            double avgPrice = Double.parseDouble(page.locator("//div[contains(text(),'Avg. NAV')]//following-sibling::div").nth(i).textContent().replaceAll("\"", "").replaceAll(",", ""));
            double currPrice = Double.parseDouble(page.locator("//span[contains(text(),'Current NAV')]//../following-sibling::div").nth(i).textContent().replaceAll("\"", "").replaceAll(",", ""));
            double pnl;
            double pnlPercent;
            mf.add(fundName);
            mf.add("MF");
            mf.add(qty);
            mf.add(avgPrice);
            invValue = qty * avgPrice;
            mf.add(formatDouble(invValue));
            mf.add(currPrice);
            currValue = qty * currPrice;
            mf.add(formatDouble(currValue));
            pnl = currValue - invValue;
            mf.add(formatDouble(pnl));
            pnlPercent = pnl / invValue;
            mf.add(formatDouble(pnlPercent));
            holdings.add(mf);
        }
    }

    private static LinkedHashMap<String, Double> getTop5Holdings() {
        LinkedHashMap<String, Double> top5 = new LinkedHashMap<>();
        for (int i = 0; i < 4; i++) {
            String key = holdings.get(i).get(0).toString();
            double value = Double.parseDouble(holdings.get(i).get(4).toString());
            top5.put(key, value);
        }
        return top5;
    }

    private static LinkedHashMap<String, Double> getTop5Gainers() {
        Collections.sort(holdings, sortByGain);
        LinkedHashMap<String, Double> top5 = new LinkedHashMap<>();
        for (int i = 0; i < 4; i++) {
            String key = holdings.get(i).get(0).toString();
            double value = Double.parseDouble(holdings.get(i).get(8).toString());
            top5.put(key, value);
        }
        return top5;
    }

    private static LinkedHashMap<String, Double> getTop5Losers() {
        Collections.sort(holdings, sortByLoss);
        LinkedHashMap<String, Double> top5 = new LinkedHashMap<>();
        for (int i = 0; i < 4; i++) {
            String key = holdings.get(i).get(0).toString();
            double value = Double.parseDouble(holdings.get(i).get(8).toString());
            top5.put(key, value);
        }
        return top5;
    }

    private static void printTop5Holdings() {
        LinkedHashMap<String, Double> top5 = getTop5Holdings();
        System.out.println(BOLD + "Top 4 Holdings\n--------------------------" + ANSI_RESET);
        outputfile.println("Top 4 Holdings\n--------------------------");
        for (Map.Entry element : top5.entrySet()) {
            String key = element.getKey().toString();
            String value = RUPEE + formatDouble((Double) element.getValue());

            System.out.format("%-15s %-15s", key, value);
            outputfile.format("%-15s %-15s", key, value);
            System.out.println();
            outputfile.println();
        }
        System.out.println(stars);
        outputfile.println(stars);
    }

    private static void printTop5Gainers() {
        LinkedHashMap<String, Double> top5 = getTop5Gainers();
        System.out.println(BOLD + "Top 4 Gainers\n--------------------------" + ANSI_RESET);
        outputfile.println("Top 4 Gainers\n--------------------------");
        for (Map.Entry element : top5.entrySet()) {
            String key = element.getKey().toString();
            String value = element.getValue().toString();
            if (value.contains("-"))
                value = "(-)" + formatDouble(Double.valueOf(value.replaceAll("-", ""))) + "%";
            else
                value = "(+)" + formatDouble(Double.valueOf(value)) + "%";
            System.out.format("%-15s %-15s", key, value);
            outputfile.format("%-15s %-15s", key, value);
            System.out.println();
            outputfile.println();
        }
        System.out.println(stars);
        outputfile.println(stars);
    }

    private static void printTop5Losers() {
        LinkedHashMap<String, Double> top5 = getTop5Losers();
        System.out.println(BOLD + "Top 4 Losers\n--------------------------" + ANSI_RESET);
        outputfile.println("Top 4 Losers\n--------------------------");
        for (Map.Entry element : top5.entrySet()) {
            String key = element.getKey().toString();
            String value = element.getValue().toString();
            if (value.contains("-"))
                value = "(-)" + formatDouble(Double.valueOf(value.replaceAll("-", ""))) + "%";
            else
                value = "(+)" + formatDouble(Double.valueOf(value)) + "%";
            System.out.format("%-15s %-15s", key, value);
            outputfile.format("%-15s %-15s", key, value);
            System.out.println();
            outputfile.println();
        }
        System.out.println(stars);
        outputfile.println(stars);
    }

    private static void fetchHoldings(String username, String password, String pin) {
        holdings = new ArrayList<>();
        Playwright playwright = Playwright.create();
        APIRequestContext requestContext = playwright.request().newContext();

        FormData form = FormData.create()
                .set("user_id", username)
                .set("password", password);

        //Login Step
        APIResponse response = requestContext.post("https://kite.zerodha.com/api/login", RequestOptions.create().setForm(form));

        if (response.status() == 200) {
            JsonObject jsonObject = new Gson().fromJson(response.text(), JsonObject.class);
            JsonObject data = jsonObject.getAsJsonObject("data");

            //Getting request_id from 1st Login Step to use it in 2nd step
            String req_id = data.get("request_id").getAsString();

            FormData form1 = FormData.create()
                    .set("twofa_value", pin)
                    .set("user_id", username)
                    .set("request_id", req_id);

            //Provide pin for the 2nd step with the request_id extracted from 1st step
            APIResponse response1 = requestContext.post("https://kite.zerodha.com/api/twofa", RequestOptions.create().setForm(form1));

            if (response1.status() == 200) {
                //From the 2nd request's response header, extract AuthToken to use it for Authentication
                Pattern p = Pattern.compile("enctoken=.+?==");
                Matcher m = p.matcher(response1.headers().toString());
                String AuthToken = null;
                if (m.find()) {

                    AuthToken = m.group(0).replaceFirst("=", " ");
                }
                Pattern p1 = Pattern.compile("public_token=.+?;");
                Matcher m1 = p1.matcher(response1.headers().toString());
                String public_token = null;
                if (m1.find()) {

                    public_token = m1.group(0).replaceFirst("public_token=", "").replaceAll(";", "");
                }

                //Query the holdings api to get holding details
                APIResponse response2 = requestContext.get("https://kite.zerodha.com/oms/portfolio/holdings",
                        RequestOptions.create().setHeader("Authorization", AuthToken)
                                .setHeader("content-type", "application/json"));

                JsonObject jsonObject2 = new Gson().fromJson(response2.text(), JsonObject.class);
                JsonArray data2 = new Gson().fromJson(jsonObject2.get("data"), JsonArray.class);

                //Query the positions api to get open position details
                APIResponse response3 = requestContext.get("https://kite.zerodha.com/oms/portfolio/positions",
                        RequestOptions.create().setHeader("Authorization", AuthToken)
                                .setHeader("content-type", "application/json"));

                JsonObject jsonObject3 = new Gson().fromJson(response3.text(), JsonObject.class);

                JsonObject data3 = new Gson().fromJson(jsonObject3.get("data"), JsonObject.class);
                JsonArray net = new Gson().fromJson(data3.get("net"), JsonArray.class);
                for (int i = 0; i < data2.size(); i++) {
                    if (data2.get(i).getAsJsonObject().get("quantity").getAsInt() + data2.get(i).getAsJsonObject().get("t1_quantity").getAsInt() != 0) {
                        addIndividualStocksToHoldingList(new Gson().fromJson(data2.get(i), JsonObject.class), "Stock");
                    } else {

                    }
                }
                for (int i = 0; i < net.size(); i++) {
                    if (net.get(i).getAsJsonObject().get("quantity").getAsInt() > 0) {
                        addIndividualStocksToHoldingList(new Gson().fromJson(net.get(i), JsonObject.class), "Stock");
                    } else {

                    }
                }
            } else {
                textArea.append("\nInvalid PIN... Please try again...");
                fl = false;
            }

        } else {
            textArea.append("\nInvalid Username or Password... Please try again...");
            fl = false;
        }
        requestContext.dispose();
        playwright.close();
    }

    private static void printHoldings() throws IOException {
        if (mutualFundFlag.equalsIgnoreCase("Y"))
            addMFHoldingsToHoldings();
        double stockInvestedVal = 0, stockCurrentVal = 0, stockGain = 0, stockGainPercent = 0;
        double mfInvestedVal = 0, mfCurrentVal = 0, mfGain = 0, mfGainPercent = 0;
        double totalInvestedVal = 0, totalCurrentVal = 0, totalGain = 0, totalGainPercent = 0;
        System.out.println(stars);
        outputfile.println(stars);
        System.out.println("Total Holdings: " + holdings.size());
        outputfile.println("Total Holdings: " + holdings.size());
        System.out.println(stars);
        outputfile.println(stars);

        String headerArray[] = {"Ticker", "Type", "Qty", "Avg Price", "Invested Val", "Curr Price", "Current Val", "P&L", "P&L %", "Weight"};
        System.out.format("%-15s %-10s %-5s %-10s %-15s %-10s %-15s %-15s %-15s %-15s", headerArray);
        outputfile.format("%-15s %-10s %-5s %-10s %-15s %-10s %-15s %-15s %-15s %-15s", headerArray);
        System.out.println();
        outputfile.println();
        System.out.println(hypens);
        outputfile.println(hypens);

        Collections.sort(holdings, sortByInvestedValue);
        int holdingSize = holdings.size();
        double totalInvested = 0;
        for (int i = 0; i < holdingSize; i++) {
            totalInvested += Double.parseDouble(holdings.get(i).get(4).toString());
        }
        for (ArrayList<Object> o : holdings) {
            //System.out.println(o);
            if (o.get(1).equals("Stock")) {
                stockInvestedVal += Double.parseDouble(o.get(4).toString());
                stockCurrentVal += Double.parseDouble(o.get(6).toString());
            } else {
                mfInvestedVal += Double.parseDouble(o.get(4).toString());
                mfCurrentVal += Double.parseDouble(o.get(6).toString());
            }
            String dataArray[] = new String[holdings.get(0).size() + 1];
            for (int i = 0; i < o.size(); i++) {
                if (i == 0 || i == 1) {
                    dataArray[i] = o.get(i).toString();
                } else if (i == 2) {
                    dataArray[i] = o.get(i).toString();
                } else if (i == (o.size() - 1)) {
                    String val = formatDouble(Double.parseDouble(o.get(i).toString()));
                    if (val.contains("-")) {
                        dataArray[i] = "(-)" + val.replaceAll("-", "") + "%";
                    } else {
                        dataArray[i] = "(+)" + val + "%";
                    }
                } else if (i == (o.size() - 2)) {
                    String val = formatDouble(Double.parseDouble(o.get(i).toString()));
                    if (val.contains("-")) {
                        dataArray[i] = "(-)" + val.replaceAll("-", "");
                    } else {
                        dataArray[i] = "(+)" + val;
                    }
                } else {
                    dataArray[i] = formatDouble(Double.parseDouble(o.get(i).toString()));
                }
                dataArray[o.size()] = formatDouble((Double.parseDouble(o.get(4).toString()) / totalInvested) * 100) + "%";
            }
            System.out.format("%-15s %-10s %-5s %-10s %-15s %-10s %-15s %-15s %-15s %-15s", dataArray);
            outputfile.format("%-15s %-10s %-5s %-10s %-15s %-10s %-15s %-15s %-15s %-15s", dataArray);
            System.out.println();
            outputfile.println();
            totalInvestedVal += Double.parseDouble(o.get(4).toString());
            //System.out.println((Double) o.get(4));
            totalCurrentVal += Double.parseDouble(o.get(6).toString());
            //System.out.println((Double) o.get(6));
        }

        System.out.println();
        outputfile.println();
        stockGain = stockCurrentVal - stockInvestedVal;
        stockGainPercent = (stockGain / stockInvestedVal) * 100;
        mfGain = mfCurrentVal - mfInvestedVal;
        mfGainPercent = (mfGain / mfInvestedVal) * 100;
        totalGain = totalCurrentVal - totalInvestedVal;
        totalGainPercent = (totalGain / totalInvestedVal) * 100;

        ArrayList<Double> detailsFromExcel = detailsfromExcel();
        if (formatDouble(totalInvestedVal).equals(formatDouble(detailsFromExcel.get(0)))
                && formatDouble(totalCurrentVal).equals(formatDouble(detailsFromExcel.get(1)))) {
            sameAsPrevHoldings = true;
        } else {
            sameAsPrevHoldings = false;
        }
        System.out.println(stars);
        outputfile.println(stars);
        System.out.println(BOLD + "Stock Stats\n--------------------------" + ANSI_RESET);
        outputfile.println("Stock Stats\n--------------------------");
        String stockStatArray[] = new String[4];
        stockStatArray[0] = "Investment: " + RUPEE + formatDouble(stockInvestedVal);
        stockStatArray[1] = "Current Value: " + RUPEE + formatDouble(stockCurrentVal);
        stockStatArray[2] = "P&L: " + RUPEE + getValueasStringWithSign(formatDouble(stockGain));
        stockStatArray[3] = "P&L Percentage: " + getValueasStringWithSign(formatDouble(stockGainPercent)) + "%";
        System.out.format("%-30s %-30s %-25s %-30s", stockStatArray);
        System.out.println();
        outputfile.format("%-30s %-30s %-25s %-30s", stockStatArray);
        outputfile.println();

        //Not able to get the Coin session id. So Mutual Fund Implementation is on hold.
        if (mutualFundFlag.equalsIgnoreCase("Y")) {
            System.out.println(stars);
            outputfile.println(stars);
            System.out.println(BOLD + "Mutual Fund Stats\n--------------------------" + ANSI_RESET);
            outputfile.println("Mutual Fund Stats\n--------------------------");
            String mfStatArray[] = new String[4];
            mfStatArray[0] = "Investment: " + RUPEE + formatDouble(mfInvestedVal);
            mfStatArray[1] = "Current Value: " + RUPEE + formatDouble(mfCurrentVal);
            mfStatArray[2] = "P&L: " + RUPEE + getValueasStringWithSign(formatDouble(mfGain));
            mfStatArray[3] = "P&L Percentage: " + getValueasStringWithSign(formatDouble(mfGainPercent)) + "%";
            System.out.format("%-30s %-30s %-25s %-30s", mfStatArray);
            System.out.println();
            outputfile.format("%-30s %-30s %-25s %-30s", mfStatArray);
            outputfile.println();
        }

        System.out.println(stars);
        outputfile.println(stars);
        System.out.println(BOLD + "Portfolio Stats\n--------------------------" + ANSI_RESET);
        outputfile.println("Portfolio Stats\n--------------------------");
        String portfolioStatArray[] = new String[4];
        portfolioStatArray[0] = "Investment: " + RUPEE + formatDouble(totalInvestedVal);
        portfolioStatArray[1] = "Current Value: " + RUPEE + formatDouble(totalCurrentVal);
        portfolioStatArray[2] = "P&L: " + RUPEE + getValueasStringWithSign(formatDouble(totalGain));
        portfolioStatArray[3] = "P&L Percentage: " + getValueasStringWithSign(formatDouble(totalGainPercent)) + "%";
        System.out.format("%-30s %-30s %-25s %-30s", portfolioStatArray);
        System.out.println();
        outputfile.format("%-30s %-30s %-25s %-30s", portfolioStatArray);
        outputfile.println();
        System.out.println(stars);
        outputfile.println(stars);
    }

    private static void addIndividualStocksToHoldingList(JsonObject jsonObject, String type) {
        String ticker;
        int quantity;
        double avg_price;
        double curr_price;
        double invested_value;
        double current_value;
        double pnl;
        double pnlPercent;

        if (type.equalsIgnoreCase("stock"))
            ticker = jsonObject.get("tradingsymbol").getAsString().replaceAll("\"", "");
        else
            ticker = jsonObject.get("fund").getAsString().replaceAll("\"", "");

        if (jsonObject.get("t1_quantity") == null) {
            if (type.equalsIgnoreCase("stock"))
                quantity = jsonObject.get("quantity").getAsInt() + jsonObject.get("overnight_quantity").getAsInt();
            else
                quantity = jsonObject.get("quantity").getAsInt();
        } else {
            quantity = jsonObject.get("quantity").getAsInt() + jsonObject.get("t1_quantity").getAsInt();
        }

        avg_price = jsonObject.get("average_price").getAsDouble();
        invested_value = quantity * avg_price;
        curr_price = jsonObject.get("last_price").getAsDouble();
        current_value = quantity * curr_price;
        pnl = current_value - invested_value;
        pnlPercent = (pnl / invested_value) * 100;

        boolean existsFlag = false;
        int indexOfMatch = 0;
        for (int i = 0; i < holdings.size(); i++) {
            if (holdings.get(i).get(0).equals(ticker)) {
                existsFlag = true;
                indexOfMatch = i;
            }
        }
        if (existsFlag) {
            quantity = (int) holdings.get(indexOfMatch).get(2) + quantity;
            invested_value = (double) holdings.get(indexOfMatch).get(4) + invested_value;
            avg_price = invested_value / quantity;
            current_value = quantity * curr_price;
            pnl = current_value - invested_value;
            pnlPercent = (pnl / invested_value) * 100;

            holdings.get(indexOfMatch).set(2, quantity);
            holdings.get(indexOfMatch).set(3, avg_price);
            holdings.get(indexOfMatch).set(4, invested_value);
            holdings.get(indexOfMatch).set(5, curr_price);
            holdings.get(indexOfMatch).set(6, current_value);
            holdings.get(indexOfMatch).set(7, pnl);
            holdings.get(indexOfMatch).set(8, pnlPercent);
        } else {
            ArrayList<Object> individual = new ArrayList<>();
            individual.add(ticker);
            individual.add(type);
            individual.add(quantity);
            individual.add(avg_price);
            individual.add(invested_value);
            individual.add(curr_price);
            individual.add(current_value);
            individual.add(pnl);
            individual.add(pnlPercent);
            holdings.add(individual);
        }
    }

    private static String formatDouble(Double value) {
        String formattedValue;
        formattedValue = String.format("%.2f", value);
        return formattedValue;
    }

    public static void writeToExcel() throws IOException {
        String filename;

        if(new File("Finance Master.xlsm").exists())
            filename = "Finance Master.xlsm";
        else
            filename = "Finance Master.xlsx";

        String sheetName = "Dashboard";
        if (!new File(filename).exists()) {
            System.out.println("\nCouldn't find an excel file named 'Finance Master.xlsx' or 'Finance Master.xlsm'");
            textArea.append("\nCouldn't find an excel file named 'Finance Master.xlsx' or 'Finance Master.xlsm'. Please make sure it exists...");
            JOptionPane.showMessageDialog(frame,
                    "Couldn't find an excel file named 'Finance Master.xlsx' or 'Finance Master.xlsm'. Please make sure it exists...",
                    "Error Information",
                    JOptionPane.ERROR_MESSAGE);
        } else {
            FileInputStream fis = new FileInputStream(filename);
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            if (wb.getSheet(sheetName) == null) {
                wb.createSheet(sheetName);
                setupDashBoardPage(wb, wb.getSheet(sheetName));
            }
            XSSFSheet sheet = wb.getSheet(sheetName);
            Collections.sort(holdings, sortByInvestedValue);
            int holdingSize = holdings.size();
            double totalInvested = 0;
            for (int i = 0; i < holdingSize; i++) {
                totalInvested += Double.parseDouble(holdings.get(i).get(4).toString());
            }
            for (int i = 0; i < holdingSize; i++) {
                int rownum = i + 3;
                //Create Row if not available
                if (sheet.getRow(rownum) == null)
                    sheet.createRow(rownum);
                //Create Columns if not available
                for (int j = 1; j <= 10; j++) {
                    if (sheet.getRow(rownum).getCell(j) == null)
                        sheet.getRow(rownum).createCell(j);
                }
                //set Column Style

                sheet.getRow(rownum).getCell(1).setCellValue(holdings.get(i).get(0).toString());
                sheet.getRow(rownum).getCell(2).setCellValue(holdings.get(i).get(1).toString());
                sheet.getRow(rownum).getCell(3).setCellValue(Double.parseDouble(holdings.get(i).get(2).toString()));
                sheet.getRow(rownum).getCell(4).setCellValue(Double.parseDouble(holdings.get(i).get(3).toString()));
                sheet.getRow(rownum).getCell(5).setCellValue(Double.parseDouble(holdings.get(i).get(4).toString()));
                sheet.getRow(rownum).getCell(6).setCellValue(Double.parseDouble(holdings.get(i).get(5).toString()));
                sheet.getRow(rownum).getCell(7).setCellValue(Double.parseDouble(holdings.get(i).get(6).toString()));
                sheet.getRow(rownum).getCell(8).setCellValue(Double.parseDouble(holdings.get(i).get(7).toString()));
                sheet.getRow(rownum).getCell(9).setCellValue(Double.parseDouble(holdings.get(i).get(8).toString()) / 100);
                sheet.getRow(rownum).getCell(10).setCellValue(Double.parseDouble(holdings.get(i).get(4).toString()) / totalInvested);
            }

            sheet.setColumnWidth(8, 11 * 256);
            sheet.setColumnWidth(9, 10 * 256);
            sheet.setColumnWidth(10, 10 * 256);

            LinkedHashMap<String, Double> top5Holdings = getTop5Holdings();
            LinkedHashMap<String, Double> top5Gainers = getTop5Gainers();
            LinkedHashMap<String, Double> top5Losers = getTop5Losers();

            LinkedHashMap<String, Double> stockPortfolioStats = getStockPortfolioStats();
            LinkedHashMap<String, Double> mutualFundPortfolioStats = getMutualFundPortfolioStats();
            LinkedHashMap<String, Double> overallPortfolioStats = getOverallPortfolioStats();

            int topHoldingsRowIterator = 4;
            for (Map.Entry mapElement : top5Holdings.entrySet()) {
                sheet.getRow(topHoldingsRowIterator).getCell(12).setCellValue(mapElement.getKey().toString());
                sheet.getRow(topHoldingsRowIterator).getCell(13).setCellValue(Double.parseDouble(mapElement.getValue().toString()));
                topHoldingsRowIterator++;
            }
            int topGainersRowIterator = 12;
            for (Map.Entry mapElement : top5Gainers.entrySet()) {
                sheet.getRow(topGainersRowIterator).getCell(12).setCellValue(mapElement.getKey().toString());
                sheet.getRow(topGainersRowIterator).getCell(13).setCellValue(Double.parseDouble(mapElement.getValue().toString()) / 100);
                topGainersRowIterator++;
            }
            int topLosersRowIterator = 20;
            for (Map.Entry mapElement : top5Losers.entrySet()) {
                sheet.getRow(topLosersRowIterator).getCell(12).setCellValue(mapElement.getKey().toString());
                sheet.getRow(topLosersRowIterator).getCell(13).setCellValue(Double.parseDouble(mapElement.getValue().toString()) / 100);
                topLosersRowIterator++;
            }

            int stockStatsRowIterator = 4;
            for (Map.Entry mapElement : stockPortfolioStats.entrySet()) {
                sheet.getRow(stockStatsRowIterator).getCell(15).setCellValue(mapElement.getKey().toString());
                sheet.getRow(stockStatsRowIterator).getCell(16).setCellValue(Double.parseDouble(mapElement.getValue().toString()));
                stockStatsRowIterator++;
            }
            int mutualFundStatsRowIterator = 12;
            for (Map.Entry mapElement : mutualFundPortfolioStats.entrySet()) {
                sheet.getRow(mutualFundStatsRowIterator).getCell(15).setCellValue(mapElement.getKey().toString());
                sheet.getRow(mutualFundStatsRowIterator).getCell(16).setCellValue(Double.parseDouble(mapElement.getValue().toString()));
                mutualFundStatsRowIterator++;
            }
            int overallPortfolioRowIterator = 20;
            for (Map.Entry mapElement : overallPortfolioStats.entrySet()) {
                sheet.getRow(overallPortfolioRowIterator).getCell(15).setCellValue(mapElement.getKey().toString());
                sheet.getRow(overallPortfolioRowIterator).getCell(16).setCellValue(Double.parseDouble(mapElement.getValue().toString()));
                overallPortfolioRowIterator++;
            }

            try {
                FileOutputStream fos = new FileOutputStream(filename);
                wb.write(fos);
                fos.close();
                wb.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
            wb.close();
//            try {
//                Desktop desktop = Desktop.getDesktop();
//                desktop.open(new File(filename));
//            } catch (Exception e) {
//                e.printStackTrace();
//            }
            textArea.append("\nWrote to Excel... Check '"+filename+"'");
        }
    }

    private static LinkedHashMap<String, Double> getOverallPortfolioStats() {
        LinkedHashMap<String, Double> myMap = new LinkedHashMap<>();
        double investmentValue = 0;
        double currentValue = 0;
        double pnl;
        double pnlPercent;
        for (ArrayList<Object> o : holdings) {
            investmentValue += Double.parseDouble(o.get(4).toString());
            currentValue += Double.parseDouble(o.get(6).toString());

        }
        pnl = currentValue - investmentValue;
        pnlPercent = pnl / investmentValue;

        myMap.put("Invested Value", investmentValue);
        myMap.put("Current Value", currentValue);
        myMap.put("P&L", pnl);
        myMap.put("% P&L", pnlPercent);
        return myMap;
    }

    private static LinkedHashMap<String, Double> getMutualFundPortfolioStats() {
        LinkedHashMap<String, Double> myMap = new LinkedHashMap<>();
        double investmentValue = 0;
        double currentValue = 0;
        double pnl;
        double pnlPercent;
        for (ArrayList<Object> o : holdings) {
            if (o.get(1).toString().equalsIgnoreCase("MF")) {
                investmentValue += Double.parseDouble(o.get(4).toString());
                currentValue += Double.parseDouble(o.get(6).toString());
            }

        }
        pnl = currentValue - investmentValue;
        pnlPercent = pnl / investmentValue;

        myMap.put("Invested Value", investmentValue);
        myMap.put("Current Value", currentValue);
        myMap.put("P&L", pnl);
        myMap.put("% P&L", pnlPercent);
        return myMap;
    }

    private static LinkedHashMap<String, Double> getStockPortfolioStats() {
        LinkedHashMap<String, Double> myMap = new LinkedHashMap<>();
        double investmentValue = 0;
        double currentValue = 0;
        double pnl;
        double pnlPercent;
        for (ArrayList<Object> o : holdings) {
            if (o.get(1).toString().equalsIgnoreCase("Stock")) {
                investmentValue += Double.parseDouble(o.get(4).toString());
                currentValue += Double.parseDouble(o.get(6).toString());
            }

        }
        pnl = currentValue - investmentValue;
        pnlPercent = pnl / investmentValue;

        myMap.put("Invested Value", investmentValue);
        myMap.put("Current Value", currentValue);
        myMap.put("P&L", pnl);
        myMap.put("% P&L", pnlPercent);
        return myMap;
    }

    private static String getValueasStringWithSign(String value) {
        String val;
        if (value.contains("-")) {
            val = "(-)" + value.replaceAll("-", "");
        } else {
            val = "(+)" + value;
        }
        return val;
    }

    private static ArrayList<Double> detailsfromExcel() throws IOException {
        String filename = "Finance Master.xlsm";
        String filename2 = "Finance Master.xlsx";
        String sheetName = "Dashboard";
        ArrayList<Double> arrayList = new ArrayList<>();
        if (!new File(filename).exists() && !new File(filename2).exists()) {
            arrayList.add(1.0);
            arrayList.add(1.0);
        } else {
            FileInputStream fis;
            try {
                fis = new FileInputStream(filename);
            } catch (Exception e) {
                fis = new FileInputStream(filename2);
            }
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            if (wb.getSheet(sheetName) == null) {
                arrayList.add(1.0);
                arrayList.add(1.0);
            } else {
                XSSFSheet sheet = wb.getSheet(sheetName);
                double investedValue = sheet.getRow(20).getCell(16).getNumericCellValue();
                double currentValue = sheet.getRow(21).getCell(16).getNumericCellValue();
                arrayList.add(investedValue);
                arrayList.add(currentValue);
            }
        }
        return arrayList;
    }
}