package pages;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.IOException;

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
    }
}
