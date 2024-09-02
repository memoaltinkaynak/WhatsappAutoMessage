# WhatsappAutoMessage

WhatsappAutoMessage, bir Excel dosyasındaki telefon numaralarına otomatik olarak mesaj gönderen bir C# konsol uygulamasıdır. Bu uygulama, WhatsApp Web kullanarak belirli bir mesajı verilen numaralara gönderir. Mesaj gönderimi sırasında manuel müdahale gerekmemesi için Selenium WebDriver kullanır.

## Gereksinimler

- [.NET SDK](https://dotnet.microsoft.com/download) (6.0 veya daha yeni bir sürüm)
- [Google Chrome](https://www.google.com/chrome/) (Güncel sürüm)
- [Chrome WebDriver](https://sites.google.com/chromium.org/driver/) (Google Chrome ile uyumlu sürüm)
- `ExcelDataReader` ve `Selenium.WebDriver` NuGet paketleri

## Kurulum

1. **Chrome WebDriver'ı İndirin**:
   - [Chrome WebDriver](https://sites.google.com/chromium.org/driver/) web sitesine gidin ve Chrome sürümünüzle uyumlu olan sürümü indirin.
   - İndirilen dosyayı çıkarın ve `chromedriver.exe` dosyasını çalıştırılabilir dosyalarınızın bulunduğu dizine ekleyin veya PATH'e ekleyin.

2. **Gerekli Paketleri Yükleyin**:
   - Proje dizininde NuGet paketlerini yükleyin:
     ```bash
     dotnet add package Selenium.WebDriver
     dotnet add package ExcelDataReader
     dotnet add package Selenium.Support
     ```

3. **Projeyi Yayınlayın**:
   - Projeyi kendine yeten bir şekilde yayınlamak için aşağıdaki komutu kullanın:
     ```bash
     dotnet publish -c Release -r win-x64 --self-contained true
     ```

## Kullanım

1. **Excel Dosyasını Hazırlayın**:
   - Excel dosyanızın ilk sütununda telefon numaraları ve ikinci sütununda mesajlar olmalıdır. İlk satır başlık olarak kullanılacaktır ve program tarafından atlanacaktır.

2. **Programı Çalıştırın**:
   - Komut istemcisinde yayınlanan dizine gidin ve uygulama dosyasını çalıştırın:
     ```bash
     <Yayınlanan_Dizin>\WhatsappAutoMessage.exe
     ```
   - Excel dosyasının tam yolunu girin ve mesaj gönderme işlemini başlatmak için bir tuşa basın.

3. **WhatsApp Web'e Giriş Yapın**:
   - Program, Google Chrome tarayıcısını açarak sizi WhatsApp Web'e yönlendirecektir.
   - QR kodunu tarayarak WhatsApp hesabınıza giriş yapın.
   - Giriş yaptıktan sonra, program devam etmek için Enter tuşuna basmanızı isteyecektir.

4. **Mesajlar Gönderilecektir**:
   - Program, Excel dosyasında listelenen telefon numaralarına mesajları otomatik olarak gönderecektir.

## Notlar

- Gönderim sürecinde herhangi bir hata oluşursa, bu hata mesajı konsolda gösterilecektir.
- `stopSending` değişkeni kullanılarak gönderim işlemi durdurulabilir.

