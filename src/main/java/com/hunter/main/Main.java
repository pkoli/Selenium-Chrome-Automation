package com.hunter.main;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebDriverException;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.sikuli.script.FindFailed;
import org.sikuli.script.Pattern;
import org.sikuli.script.Screen;

import java.io.*;
import java.util.*;
import java.util.concurrent.TimeUnit;

/**
 * Created by Pavan on 29-05-2017.
 */
public class Main {

    private static Properties properties;
    static{
        try {
            properties = new Properties();
            File file = new File("essential.properties");
            FileInputStream fileInputStream = new FileInputStream(file);
            properties.load(fileInputStream);
        } catch (IOException e) {
            System.out.println("Couldn't load properties file, exiting");
        }
    }

    public static void main(String[] args) throws IOException, FindFailed {
        ChromeDriver driver = getChromeDriver();

        Screen screen = new Screen();
        Pattern extensionClick = new Pattern("images/email-hunter.PNG");
        Pattern exportAllClick = new Pattern("images/exportAll.PNG");
        Pattern disableAllExport = new Pattern("images/disableExportAll.PNG");


        File inputFile = new File("Input.txt");
        FileInputStream inputFileInputStream = new FileInputStream(inputFile);
        Scanner inputScanner = new Scanner(inputFileInputStream);

        Map<String, List<String>> urlMailMapping = new LinkedHashMap<String, List<String>>();

        disableExportAll(driver, screen, extensionClick, disableAllExport);

        String url;
        try {
            while (inputScanner.hasNextLine()) {
                url = inputScanner.nextLine();
                try {
                    if(driver.getWindowHandle() != driver.getWindowHandles().toArray()[driver.getWindowHandles().toArray().length-1]){
                        driver.switchTo().window(String.valueOf(driver.getWindowHandles().toArray()[driver.getWindowHandles().toArray().length-1]));
                    }
                    driver.navigate().to((url.trim()));
                } catch (TimeoutException e) {
                    driver.close();
                    driver=getChromeDriver();
                    disableExportAll(driver, screen, extensionClick, disableAllExport);
                    urlMailMapping.put(url, Arrays.asList("Timed out"));
                    continue;
                }
                catch (WebDriverException e) {
                    driver.close();
                    driver=getChromeDriver();
                    disableExportAll(driver, screen, extensionClick, disableAllExport);
                    urlMailMapping.put(url, Arrays.asList("Timed out"));
                    continue;
                }

                if (findEmailsAndMapToURL(screen, extensionClick, exportAllClick, urlMailMapping, url)) continue;
            }
        } catch (Exception e) {
            System.out.println("Exception"+e);
        } finally {
            for (int i = 0; i < urlMailMapping.size(); i++) {
                createExcelSheet(urlMailMapping, urlMailMapping.keySet().toArray(new String[urlMailMapping.size()]));
            }
            driver.close();
            System.out.println("Exited Successfully, Excel file created");
        }
    }

    private static boolean findEmailsAndMapToURL(Screen screen, Pattern extensionClick, Pattern exportAllClick, Map<String, List<String>> urlMailMapping, String url) throws FindFailed, InterruptedException, IOException {
        screen.click(extensionClick);

        TimeUnit.SECONDS.sleep(3);

        try {
            screen.click(exportAllClick);
        } catch (FindFailed e) {
            screen.click(extensionClick);
            urlMailMapping.put(url, Arrays.asList("No email found"));
            TimeUnit.SECONDS.sleep(1);
            return true;
        }

        TimeUnit.SECONDS.sleep(1);

        screen.click(extensionClick);
        try {
            File allCollectedEmails = new File(properties.getProperty("downloadPath") + "\\CurrentPageList.txt");

            FileInputStream fileInputStream = new FileInputStream(allCollectedEmails);


            Scanner sc = new Scanner(fileInputStream);

            List<String> emails = new ArrayList<String>();

            while (sc.hasNextLine()) {
                String email = sc.nextLine();
                emails.add(email);
            }

            sc.close();
            fileInputStream.close();

            System.out.println(allCollectedEmails.delete());

            urlMailMapping.put(url, emails);
            System.out.println("Mapped values" + url + " " + emails);
        }
        catch (FileNotFoundException e){
            screen.click(extensionClick);
            urlMailMapping.put(url, Arrays.asList("No email found"));
            TimeUnit.SECONDS.sleep(1);
            return true;
        }
        return false;
    }

    private static void disableExportAll(ChromeDriver driver, Screen screen, Pattern extensionClick, Pattern disableAllExport) throws FindFailed {
        driver.get("http://www.google.com");

        screen.click(extensionClick);
        screen.click(disableAllExport);
        screen.click(extensionClick);
    }

    private static ChromeDriver getChromeDriver() {
        ChromeOptions options = new ChromeOptions();

        options.addExtensions(new File("hunterCore.crx"));
        options.addArguments("test-type");
        options.addArguments("disable-popup-blocking");

        System.setProperty("webdriver.chrome.driver", "chromedriver.exe");

        DesiredCapabilities capabilities = DesiredCapabilities.chrome();
        capabilities.setCapability(ChromeOptions.CAPABILITY, options);
        ChromeDriver chromeDriver = new ChromeDriver(capabilities);
        chromeDriver.manage().timeouts().pageLoadTimeout(Long.parseLong(properties.getProperty("timeoutSeconds")),TimeUnit.SECONDS);
        return chromeDriver;
    }

    private static void createExcelSheet(Map<String, List<String>> urls, String[] URLS) {
        Object [][] mapping = new Object[URLS.length][2];

        for (int i=0; i< URLS.length; i++){
            if (!urls.get(URLS[i]).isEmpty()) {
                String foundEmails = urls.get(URLS[i]).toString();
                mapping[i][0] = URLS[i];
                mapping[i][1] = foundEmails.substring(1, foundEmails.length()-1);
            }
            else {
                mapping[i][0] = URLS[i];
                mapping[i][1] = "No email found";
            }
        }


        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("URL-Email");

        int rowCount = 0;

        for (Object[] aBook : mapping) {
            Row row = sheet.createRow(++rowCount);

            int columnCount = 0;

            for (Object field : aBook) {
                Cell cell = row.createCell(++columnCount);
                if (field instanceof String) {
                    cell.setCellValue((String) field);
                } else if (field instanceof Integer) {
                    cell.setCellValue((Integer) field);
                }
            }

        }

        FileOutputStream outputStream = null;
        try {
            outputStream = new FileOutputStream("Hunter_catch.xlsx");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        try {
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}