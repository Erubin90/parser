package parser;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Connection;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.stream.Collectors;

public class Main {
    public static void main(String[] args) {
        WebDriver driver = getChromeDriver();
        Scanner scanner = new Scanner(System.in);
        System.setProperty("webdriver.chrome.driver", "src/main/resources/chromedriver");
        String url = getUrl(scanner);
        String pathSaveFile = getPathSaveFile(scanner);
        Connection connection = Jsoup.connect(url);
        String valueCity = getValueCity(scanner, connection);
        long timeStart = System.currentTimeMillis();
        driver.get(url);
        List<List<String>> infoAllPage = getParsing(driver, valueCity);
        driver.close();
        putExel(infoAllPage, pathSaveFile);
        System.out.println("Времы работы программы " + (System.currentTimeMillis() - timeStart) / 1000.0);
    }

    public static String getUrl(Scanner scanner){
        boolean flag = true;
        System.out.print("Введите URL сайта catalogue.hyve.ru\n▶ ");
        String url = scanner.nextLine();
        while (flag) {
            if (!url.matches("https://catalogue.hyve.ru/\\w+-\\w+/exhibitorlist.aspx.project_id=\\d+")){
                System.out.print("Введен не верный url\n▶ ");
                url = scanner.nextLine();
            }
            else {
                flag = false;
            }
        }
        return url;
    }

    public static String getPathSaveFile(Scanner scanner){
        System.out.print("Укажите путь куда сохранить файл и название файла\n" +
                "Пример на разных ОС\n" +
                "Windows - C:\\Users\\Имя_Пользователя\\Desktop\\parser \n" +
                "Mac /Users/Имя_Пользователя/Desktop/parser\n" +
                "Linux ты и так знаешь что сюда подставлять)\n" +
                "▶ ");
        String path = scanner.nextLine();
        return path + ".xlsx";
    }

    public static String getValueCity(Scanner scanner, Connection connection) {
        System.out.print("Фильтровать по городам?\nДа Нет\n▶ ");
        String filterFlag = scanner.nextLine();
        if (filterFlag.equals("Да") || filterFlag.equals("да")) {
            List<String> cityList = getCityList(connection);
            if (!cityList.isEmpty()) {
                for (String str : cityList) {
                    System.out.println(str);
                }
                boolean flag = true;
                System.out.println("Выбирете город для фильтра. Чтоб выбрать все города отправьте -");
                String city = "";
                while (flag) {
                    city = scanner.nextLine();
                    flag = !cityList.contains(city);
                    if (flag) {
                        System.out.println("Город добавлен в фильтр");
                    } else {
                        System.out.print("Скопируйте название города который вам нужен\n▶ ");
                    }
                }
                return city;
            }
            else {
                return "-";
            }
        }
        else {
            return "-";
        }
    }

    public static List<String> getCityList(Connection connection) {
        List<String> list = new ArrayList<>();
        Elements elements = null;
        for (int i = 1; i <= 5; i++ ) {
            try {
                elements = Objects.requireNonNull(connection
                        .get()
                        .getElementById("p_lt_zoneContainer_pageplaceholder_p_lt_zoneForm_Filter_filterControl_ddlCountries"))
                        .getElementsByAttribute("value");
                break;
            } catch (IOException e) {
                if (i < 5) {
                    System.out.println("Ошибка: " + e.getMessage());
                }
                else {
                    System.out.println("Завершение программы из-за ошибки: " + e.getMessage());
                }
            }
        }
        if (elements != null) {
            for (Element element : elements) {
                list.add(element.text());
            }
        }
        return list;
    }

    public static WebDriver getChromeDriver() {
        ChromeOptions chromeOptions = new ChromeOptions();
        chromeOptions.addArguments("headless");
        return new ChromeDriver(chromeOptions);
//        return new ChromeDriver();
    }

    public static List<List<String>> getParsing(WebDriver driver, String valueCity) {
        System.out.println("Парсинг данный начался");
        List<WebElement> elements = driver
                .findElement(By.id("p_lt_zoneContainer_pageplaceholder_p_lt_zoneForm_Filter_filterControl_ddlCountries"))
                .findElements(By.tagName("option"));
        for (WebElement el : elements) {
            if (el.getText().equals(valueCity)) {
                el.click();
            }
        }

        driver.findElement(By.className("FilterSubmit")).click();
        int maxCountPage = getMaxCountPage(driver);
        List<List<String>> allDate = new ArrayList<>();
        allDate.add(List.of("name", "Country", "Pavilion", "Stand", "Address",
                "Telephone number", "Website", "Email", "About company", "Brands"));

        for (int i = 1; i <= maxCountPage; i++) {
            var allLinePage = driver.findElement(By.className("exhibitor_list"))
                    .findElements(By.tagName("a"));
            for (WebElement linePage : allLinePage) {
                List<String> date = new ArrayList<>();
                for (WebElement wb : linePage.findElements(By.tagName("div"))) {
                    date.add(wb.getText());
                }
                String url = linePage.getAttribute("href");
                if (!url.matches(".*#")) {
                    date.addAll(getInfoCompany(url));
                }
                allDate.add(date);
            }
            List<WebElement> pager = getPager(driver);
            System.out.println("[+] Страница " + i + " из " + maxCountPage);
            pager.get(pager.size() - 2).click();
        }
        System.out.println("Парсинг данный завершен");
        return allDate;
    }

    public static int getMaxCountPage(WebDriver driver) {
        List<String> allPageSelectors = getPager(driver)
                .stream()
                .map(x -> x.getAttribute("href"))
                .collect(Collectors.toList());
        return Integer.parseInt(allPageSelectors.get(allPageSelectors.size() - 1)
                .replaceAll("\\D", ""));
    }

    public static List<WebElement> getPager (WebDriver driver) {
        return driver.findElement(By.className("pager"))
                .findElements(By.tagName("a"));
    }

    public static List<String> getInfoCompany(String url){
        List<String> data = new ArrayList<>();
        for (int j = 1; j <= 5; j++) {
            try {
                var connection = Jsoup.connect(url)
                        .get()
                        .getElementsByClass("scorecard")
                        .get(0)
                        .getElementsByTag("div");
                for (int i = 0; i < connection.size(); i++) {
                    switch (i) {
                        case 3:
                        case 6:
                        case 8:
                        case 10:
                        case 12:
                        case 13:
                            data.add(connection.get(i).text());
                            break;
                    }
                }
                break;
            } catch (IOException e) {
                System.out.println("Неудалось подключиться к " + url);
            }
        }
        return data;
    }

    public static void putExel(List<List<String>> infoAllPage, String path) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("parse");

        System.out.println("Создание excel файла");
        int rowNum = 0;
        for (List<String> page : infoAllPage) {
            Row row = sheet.createRow(rowNum++);
            int colNum = 0;
            for (String str : page) {
                Cell cell = row.createCell(colNum++);
                cell.setCellValue(str);
            }
        }
        try {
            FileOutputStream outputStream = new FileOutputStream(path);
            workbook.write(outputStream);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
