using System;
using System.Data;
using System.IO;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using ExcelDataReader;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;

class Program
{
    private static ChromeDriver driver;
    private static DataTable excelData;
    private static bool stopSending = false;

    static void Main(string[] args)
    {
        Console.WriteLine("Mesaj göndermek için bir Excel dosyası seçin:");
        string filePath = Console.ReadLine();

        if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
        {
            Console.WriteLine("Geçerli bir dosya yolu girin.");
            return;
        }

        LoadExcel(filePath);

        if (excelData != null)
        {
            Console.WriteLine("Excel dosyası yüklendi. Mesaj gönderme işlemini başlatmak için herhangi bir tuşa basın...");
            Console.ReadKey();
            SendMessages();
        }

        Console.WriteLine("İşlem tamamlandı.");
    }

    private static void LoadExcel(string path)
    {
        try
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    excelData = result.Tables[0];
                }
            }
            Console.WriteLine("Excel dosyası başarıyla yüklendi.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Excel dosyası yüklenirken bir hata oluştu: {ex.Message}");
        }
    }

    private static void SendMessages()
    {
        if (excelData == null)
        {
            Console.WriteLine("Lütfen önce bir Excel dosyası seçin.");
            return;
        }

        stopSending = false;

        try
        {
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--start-maximized");
            driver = new ChromeDriver(options);
            driver.Navigate().GoToUrl("https://web.whatsapp.com");

            Console.WriteLine("Lütfen WhatsApp Web'e giriş yapın ve devam etmek için Enter tuşuna basın...");
            Console.ReadLine();

            // İlk satırı atlamak için sayaç oluştur
            bool isFirstRow = true;

            foreach (DataRow row in excelData.Rows)
            {
                if (stopSending)
                    break;

                // İlk satırı atla
                if (isFirstRow)
                {
                    isFirstRow = false;
                    continue;
                }

                string phoneNumber = row[0].ToString();
                string message = row[1].ToString();

                if (string.IsNullOrWhiteSpace(phoneNumber))
                    break;

                SendMessage(phoneNumber, message);
                Console.WriteLine($"Mesaj gönderildi: {phoneNumber}");
            }

            Console.WriteLine("Listenizdeki tüm adreslere mesaj gönderildi.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Mesaj gönderimi sırasında bir hata oluştu: {ex.Message}");
        }
        finally
        {
            driver?.Quit();
        }
    }

    private static void SendMessage(string phoneNumber, string message)
    {
        try
        {
            driver.Navigate().GoToUrl($"https://web.whatsapp.com/send?phone={phoneNumber}&text={Uri.EscapeDataString(message)}");

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//span[@data-icon='send']"))).Click();

            System.Threading.Thread.Sleep(2000); // Mesajın gönderilmesi için kısa bir bekleme süresi
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Telefon numarası: {phoneNumber} için mesaj gönderilemedi: {ex.Message}");
        }
    }
}
