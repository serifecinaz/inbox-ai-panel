/**
 * ==============================================================================
 * GMAIL AKILLI AYIKLAMA VE ETİKETLEME OTOMASYONU (Google Apps Script)
 * Model: Gemini 2.5 Flash (Google AI Studio)
 * ==============================================================================
 */

// ========================== YAPILANDIRMA (AYARLAR) ==========================
const CONFIG = {
  // Gemini API Anahtarınızı 'Script Properties' (Komut Dosyası Özellikleri) içine
  // 'GEMINI_API_KEY' adıyla eklemeniz önerilir. İsterseniz tırnak içine doğrudan da yazabilirsiniz:
  API_KEY: PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || 'AQ.Ab8RN6JqLxQjilhmOCmz4xip7jHR_pk2GUxYC0zbvJzSWd6erg',
  
  // Kullanılacak Gemini modeli (Hızlı, doğru ve ücretsiz kota dostu)
  GEMINI_MODEL: 'gemini-3.8-flash',

  // Etiket Ayarları:
  LABEL_PREFIX: 'AI/',           // Etiketlerin başına gelecek ön ek (örn: AI/Finans, AI/Satış)
  PROCESSED_LABEL: 'AI/İşlendi', // Tekrar tekrar işlenmeyi engelleyen işaret etiketi

  // Tek seferde taranacak maksimum e-posta sayısı (Zaman aşımını önlemek için)
  MAX_EMAILS_PER_RUN: 15,

  // YÖNLENDİRME (İleride aktif etmek istediğinizde true yapmanız yeterlidir)
  ENABLE_FORWARDING: false,
  FORWARDING_RULES: {
    'Finans-Fatura': 'muhasebe@sirketiniz.com',
    'Satış-Pazarlama': 'satis@sirketiniz.com',
    'Destek-Talep': 'destek@sirketiniz.com',
    'İnsan Kaynakları': 'ik@sirketiniz.com'
  },

  // E-postalar işlendikten sonra size tek bir sabah özeti gönderilsin mi?
  SEND_DIGEST_EMAIL: false, // true yaparsanız kendi adresinize özet maili atar
};

// ============================== ANA ÇALIŞMA FONKSİYONU ==============================
/**
 * Zamanlayıcı (Trigger) tarafından her sabah veya periyodik olarak çalıştırılacak fonksiyon.
 */
function processInboxEmails() {
  const apiKey = CONFIG.API_KEY;
  if (!apiKey || apiKey === 'BURAYA_GEMINI_API_KEY_GIRINIZ') {
    Logger.log('HATA: Lütfen geçerli bir Gemini API Anahtarı tanımlayın!');
    return;
  }

  // Sadece gelen kutusunda olan, okunmamış ve henüz AI tarafından işlenmemiş mailleri ara
  const query = 'in:inbox is:unread -label:' + CONFIG.PROCESSED_LABEL;
  const threads = GmailApp.search(query, 0, CONFIG.MAX_EMAILS_PER_RUN);

  Logger.log('İşlenecek e-posta zinciri sayısı: ' + threads.length);
  if (threads.length === 0) {
    Logger.log('İşlenecek yeni e-posta bulunamadı.');
    return;
  }

  const processedLabel = getOrCreateLabel(CONFIG.PROCESSED_LABEL);
  const summaryReport = [];

  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    const messages = thread.getMessages();
    const latestMessage = messages[messages.length - 1]; // En güncel mesajı al

    const sender = latestMessage.getFrom();
    const subject = latestMessage.getSubject() || '(Konusuz)';
    const plainBody = latestMessage.getPlainBody();
    // Gemini'a göndermek için içeriği makul bir uzunlukta kes (ilk 2000 karakter yeterlidir)
    const snippet = plainBody.substring(0, 2000);

    Logger.log(`[${i + 1}/${threads.length}] Analiz ediliyor: "${subject}" | Kimden: ${sender}`);

    try {
      // 1. Gemini ile e-postayı sınıflandır
      const aiResult = classifyEmailWithGemini(subject, sender, snippet, apiKey);

      if (aiResult) {
        // 2. Kategori etiketini uygula (Örn: AI/Finans-Fatura)
        const categoryLabelName = CONFIG.LABEL_PREFIX + aiResult.category;
        const categoryLabel = getOrCreateLabel(categoryLabelName);
        thread.addLabel(categoryLabel);

        // 3. Aciliyet etiketi uygula (Örn: AI/Acil)
        if (aiResult.urgency === 'Acil') {
          const urgentLabel = getOrCreateLabel(CONFIG.LABEL_PREFIX + 'Acil');
          thread.addLabel(urgentLabel);
        }

        // 4. İşlendi etiketini ekle (Bir daha taranmasın)
        thread.addLabel(processedLabel);

        // 5. İleride yönlendirme açılırsa:
        if (CONFIG.ENABLE_FORWARDING && CONFIG.FORWARDING_RULES[aiResult.category]) {
          const targetAddress = CONFIG.FORWARDING_RULES[aiResult.category];
          Logger.log(`Yönlendiriliyor: ${targetAddress}`);
          latestMessage.forward(targetAddress, {
            replyTo: latestMessage.getFrom(),
            subject: `[Yönlendirildi: ${aiResult.category}] ` + subject
          });
        }

        summaryReport.push({
          subject: subject,
          sender: sender,
          category: aiResult.category,
          urgency: aiResult.urgency,
          summary: aiResult.summary
        });
      }
    } catch (err) {
      Logger.log('E-posta işlenirken hata oluştu: ' + err.toString());
    }
  }

  // İsteğe bağlı özet e-postası gönder
  if (CONFIG.SEND_DIGEST_EMAIL && summaryReport.length > 0) {
    sendMorningDigest(summaryReport);
  }

  Logger.log('Tüm mailler başarıyla işlendi.');
}

// ============================== GEMINI API ENTEGRASYONU ==============================
/**
 * Gemini 2.5 Flash modeline e-postayı analiz ettirip JSON çıktısı alır.
 */
function classifyEmailWithGemini(subject, sender, body, apiKey) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.GEMINI_MODEL}:generateContent?key=${apiKey}`;

  const prompt = `
Sen bir kurumsal e-posta asistanısın. Görevin gelen e-postayı analiz edip konusuna göre kategorize etmek, aciliyetini belirlemek ve çok kısa bir özet çıkarmaktır.

E-POSTA BİLGİLERİ:
- Kimden: ${sender}
- Konu: ${subject}
- İçerik Özeti:
${body}

Lütfen e-postayı aşağıdaki KATEGORİLERDEN YALNIZCA BİRİNE ata:
- "Finans-Fatura" (Fatura, dekont, ödeme bildirimi, banka mailleri, muhasebe)
- "Satış-Pazarlama" (Yeni teklif talepleri, potansiyel müşteriler, ticari iş birlikleri)
- "İş-Müşteri" (Mevcut projeler, toplantı davetleri, iş takibi, müşteri talepleri)
- "Destek-Talep" (Teknik problemler, hata bildirimleri, yardım talepleri)
- "Bülten-Reklam" (Toplu bültenler, abonelikler, promosyon ve kampanya mailleri)
- "Kişisel" (Kişisel iletişimler, özel bildirimler)
- "Diğer" (Yukarıdakilere uymayan genel mailler)

ACİLİYET SEVİYELERİ:
- "Acil" (Bugün içinde aksiyon alınması gereken, kriz, kritik uyarı vb.)
- "Normal" (Rutin işler)
- "Düşük" (Bültenler, bilgilendirmeler)

Lütfen cevabını SADECE aşağıdaki JSON formatında ver, fazladan markdown veya açıklama yazma:
{
  "category": "Kategori İsmi",
  "urgency": "Acil | Normal | Düşük",
  "summary": "1-2 cümlelik Türkçe kısa özet"
}
`;

  const payload = {
    contents: [
      {
        parts: [{ text: prompt }]
      }
    ],
    generationConfig: {
      responseMimeType: "application/json",
      temperature: 0.1
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();

  if (responseCode !== 200) {
    Logger.log('Gemini API Hatası (' + responseCode + '): ' + response.getContentText());
    return null;
  }

  const jsonResponse = JSON.parse(response.getContentText());
  const aiText = jsonResponse.candidates[0].content.parts[0].text;
  return JSON.parse(aiText);
}

// ============================== YARDIMCI FONKSİYONLAR ==============================
/**
 * Gmail etiketi varsa getirir, yoksa otomatik oluşturur.
 */
function getOrCreateLabel(name) {
  let label = GmailApp.getUserLabelByName(name);
  if (!label) {
    label = GmailApp.createLabel(name);
  }
  return label;
}

/**
 * Bağlantıyı ve API Key'i hızlıca test etmek için tek seferlik test fonksiyonu.
 */
function testGeminiConnection() {
  Logger.log('Gemini API bağlantısı test ediliyor...');
  const test = classifyEmailWithGemini(
    'Fatura Hk. - Mart 2026',
    'muhasebe@ornekfirma.com',
    'Sayın Yetkili, Mart ayına ait 15.000 TL tutarındaki faturanız ektedir. İyi çalışmalar.',
    CONFIG.API_KEY
  );
  Logger.log('Test Sonucu: ' + JSON.stringify(test, null, 2));
}

/**
 * Sabah işlenen maillerin özetini kendi adresinize gönderme fonksiyonu.
 */
function sendMorningDigest(summaryReport) {
  const myEmail = Session.getActiveUser().getEmail();
  let body = 'Günaydın! Sabah gelen kutunuz yapay zeka tarafından ayıklandı:\n\n';

  summaryReport.forEach((item, index) => {
    body += `${index + 1}. [${item.category}] [${item.urgency}] ${item.subject}\n`;
    body += `   Kimden: ${item.sender}\n`;
    body += `   Özet: ${item.summary}\n\n`;
  });

  GmailApp.sendEmail(myEmail, '☀️ Sabah E-posta Özeti & Triyaj Raporu', body);
}
