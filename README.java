    import java.awt.Image;
    import java.io.File;
    import java.io.FileInputStream;
    import java.io.FileNotFoundException;
    import java.io.IOException;
    import java.util.List;

    import javax.imageio.ImageIO;

    import org.apache.poi.ss.usermodel.Row;
    import org.apache.poi.xssf.usermodel.XSSFSheet;
    import org.apache.poi.xssf.usermodel.XSSFWorkbook;
    import org.openqa.selenium.By;
    import org.openqa.selenium.TakesScreenshot;
    import org.openqa.selenium.WebDriver;
    import org.openqa.selenium.WebElement;
    import org.openqa.selenium.chrome.ChromeDriver;
    import org.testng.annotations.BeforeTest;

    import ru.yandex.qatools.ashot.AShot;
    import ru.yandex.qatools.ashot.Screenshot;
    import ru.yandex.qatools.ashot.shooting.ShootingStrategies;

    public class OTAFinalScript {

      public static void main(String[] args) throws InterruptedException, IOException{
        // TODO Auto-generated method stub


        System.setProperty("webdriver.chrome.driver", "C:\\Users\\kalynith\\Downloads\\chromedriver_win32 (1)\\chromedriver.exe");
        WebDriver driver = new ChromeDriver();
        driver.manage().window().maximize();

        driver.get("https://admin.stage.e2ro.com/");
        Thread.sleep(2000);

        driver.findElement(By.xpath("//button[text()='Login with Okta']")).click();
        Thread.sleep(15000);

        driver.findElement(By.xpath("//label[text()='Remember me']")).click();
        Thread.sleep(2000);
        driver.findElement(By.id("okta-signin-submit")).click();
        Thread.sleep(12000);

        // Enter your network name
        driver.findElement(By.xpath("//input[@name='q']")).sendKeys("OTA");
        Thread.sleep(3000);

        driver.findElement(By.xpath("//button[@type='submit']")).click();
        Thread.sleep(2000);



        driver.findElement(By.xpath("//a[contains(@class,'eo-link') and contains(@href, '/networks/')]")).click();
        Thread.sleep(12000);

        driver.navigate().refresh();
        Thread.sleep(2000);

        WebElement firmwareBox = driver.findElement(By.className("ant-input"));
        Thread.sleep(2000);

        //Importing file data
            File file = new File("C:\\Users\\kalynith\\Documents\\Firmware_Data.xlsx");
            FileInputStream fileName = new FileInputStream(file);

            //Creating workbook sheet
            XSSFWorkbook workbook = new XSSFWorkbook(fileName);
            XSSFSheet sheet = workbook.getSheetAt(0);
            int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();

            //Loop overall row in the excel sheet
                for(int i=1; i< rowCount+1; i++) {
                  String inputData = sheet.getRow(i).getCell(0).getStringCellValue();
                  String outpuData = sheet.getRow(i).getCell(1).getStringCellValue();
                firmwareBox.sendKeys(inputData);
                Thread.sleep(3000);

                WebElement updateButton = driver.findElement(By.xpath("//button[@type = 'submit' and text()='Update'][1]"));
                updateButton.click();
                Thread.sleep(3000);

                WebElement refreshButton = driver.findElement(By.xpath("//span[text()='Refresh']//parent::button"));
                refreshButton.click();
                Thread.sleep(30000);
                refreshButton.click();

                System.out.println(inputData + " OTA Update result");

                List<WebElement> secondaryFirmwarText = driver.findElements(By.xpath("//p[contains(@class,'gr-2@xl')][7]//span//span"));

                for(int j=0; j<secondaryFirmwarText.size(); j++) {
                  String firmwareName = secondaryFirmwarText.get(j).getText();
                  System.out.println("Node " +(j+1)+ ":");

                  if(firmwareName.contains(outpuData)) {
                    System.out.println("Firmware " + inputData +  " is updated");			
                  }else {
                    System.out.println("Firmware " + inputData + " is not updated");
                  }				
                }
                Screenshot scrShot = new AShot().shootingStrategy(ShootingStrategies.viewportPasting(ShootingStrategies.scaling(1.75f), 2000)).takeScreenshot(driver);
                ImageIO.write(scrShot.getImage(),"png", new File("C:\\Users\\kalynith\\Documents\\Screenshot\\fullimage "+outpuData+".png"));
                System.out.println(" ");
                       workbook.close();
            }		
           driver.close();
        }

    }
