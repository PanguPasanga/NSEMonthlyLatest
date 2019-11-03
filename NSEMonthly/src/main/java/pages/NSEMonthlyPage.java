package pages;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import utils.MonthlyNiftyExcel;

import java.io.IOException;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

public class NSEMonthlyPage {
    public static void main(String args[]) throws IOException, InterruptedException {
        String CE_oi, CE_choi, CE_vol, CE_iv, CE_ltp, PE_ltp, PE_iv, PE_vol, PE_choi, PE_oi, livePrice, strikeData = null;


        System.setProperty("webdriver.chrome.driver", "./Driver/chromedriver.exe");
        WebDriver driver = new ChromeDriver();
        driver.navigate().to("https://www1.nseindia.com/");
        driver.manage().window().maximize();
        Actions act = new Actions(driver);
        WebDriverWait wt = new WebDriverWait(driver, 20);
        WebElement live = driver.findElement(By.linkText("Live Market"));

        act.moveToElement(live).perform();

        WebElement equity = driver.findElement(By.xpath("//li[@id=\"main_livewth_oc\"]/a"));
        act.moveToElement(equity);
        act.click().build().perform();

        WebElement expirtySelect = driver.findElement(By.xpath("//select[@id=\"date\"]"));
        Select sc = new Select(expirtySelect);
        sc.selectByValue("28NOV2019");
        //Dynamic table starts here
        WebElement optionTable = driver.findElement(By.xpath("//table[@id=\"octable\"]"));
        List<String> strikePrices=new ArrayList<String>();
        strikePrices.add("10700.00");
        strikePrices.add("10800.00");
        strikePrices.add("10900.00");
        strikePrices.add("11000.00");
        strikePrices.add("11100.00");
        strikePrices.add("11200.00");
        strikePrices.add("11300.00");
        strikePrices.add("11400.00");
        strikePrices.add("11500.00");
        strikePrices.add("11600.00");
        strikePrices.add("11700.00");
        strikePrices.add("11800.00");
        strikePrices.add("11900.00");
        strikePrices.add("12000.00");
        List<String> optionData = new ArrayList<String>();
        for(String s:strikePrices)
        {
            //System.out.println(s);
            WebElement strikePrice = optionTable.findElement(By.xpath("//b[contains(text(),'"+s+"')]"));

            WebElement callOpenInterest = strikePrice.findElement(By.xpath("../parent::td/preceding-sibling::td/following-sibling::td"));
            WebElement callChangeInInterest = strikePrice.findElement(By.xpath("../parent::td/preceding-sibling::td/following-sibling::td[2]"));
            WebElement callVolume = strikePrice.findElement(By.xpath("../parent::td/preceding-sibling::td/following-sibling::td[3]"));
            WebElement callIV = strikePrice.findElement(By.xpath("../parent::td/preceding-sibling::td/following-sibling::td[4]"));
            WebElement callLTP=strikePrice.findElement(By.xpath("../parent::td/preceding-sibling::td/following-sibling::td[5]"));
            WebElement putLTP=strikePrice.findElement(By.xpath("../parent::td/preceding-sibling::td/following-sibling::td[17]"));
            WebElement putIV = strikePrice.findElement(By.xpath("../parent::td/preceding-sibling::td/following-sibling::td[18]"));
            WebElement putVolume = strikePrice.findElement(By.xpath("../parent::td/preceding-sibling::td/following-sibling::td[19]"));
            WebElement putChangeInInterest=strikePrice.findElement(By.xpath("../parent::td/preceding-sibling::td/following-sibling::td[20]"));
            WebElement putOpenInterest =strikePrice.findElement(By.xpath("../parent::td/preceding-sibling::td/following-sibling::td[21]"));
            WebElement niftyPrice=driver.findElement(By.xpath("//nobr[contains(text(),1)]"));


            //Converted webelement to string
            CE_oi=callOpenInterest.getText();
            CE_choi=callChangeInInterest.getText();
            CE_vol=callVolume.getText();
            CE_iv=callIV.getText();
            CE_ltp=callLTP.getText();
            PE_ltp=putLTP.getText();
            PE_iv=putIV.getText();
            PE_vol=putVolume.getText();
            PE_choi=putChangeInInterest.getText();
            PE_oi=putOpenInterest.getText();
            livePrice=niftyPrice.getText();
            strikeData=strikePrice.getText();

            DayOfWeek day= LocalDate.now().getDayOfWeek();
            LocalDate date=LocalDate.now();
            String date_day=date.toString()+day.toString();
            System.out.println(date_day);

            //Adding into the list
            optionData.add(date_day);
            optionData.add(CE_oi);
            optionData.add(CE_choi);
            optionData.add(CE_vol);
            optionData.add(CE_iv);
            optionData.add(CE_ltp);
            optionData.add(PE_ltp);
            optionData.add(PE_iv);
            optionData.add(PE_vol);
            optionData.add(PE_choi);
            optionData.add(PE_oi);
            optionData.add(livePrice);

            for(String b:optionData) {
                System.out.println(b);

            }
            System.out.println("------------------End of strike-----------------");
            MonthlyNiftyExcel.writeToExcel(optionData,strikeData);
            optionData.clear();
        }
    }
}
