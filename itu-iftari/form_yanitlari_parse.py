import os
import secrets
import string
import openpyxl


# Rastgele 6 haneli eşsiz kod üretme fonksiyonu
def generate_unique_code(existing_codes):
    while True:
        code = ''.join(secrets.choice(string.ascii_letters + string.digits) for _ in range(6))
        if code not in existing_codes:
            return code


def create_davetli_listesi_js(veriler):
    # JavaScript dosyasını oluştur
    js_icerik = "davetliler = {\n"
    js_icerik += ",\n".join(f'  "{veri[TOKEN]}": "{veri[ISIM]}"' for veri in veriler)
    js_icerik += "\n}\n"

    with open("davetli-listesi.js", "w", encoding="utf-8") as js_dosyasi:
        js_dosyasi.write(js_icerik)


def read_and_parse_form_yanitlari_xlsx(dosya_adi):
    # Excel dosyasını yükle
    wb = openpyxl.load_workbook(dosya_adi, data_only=True)
    ws = wb.active  # İlk sayfayı seç

    # Sütun başlıklarını al
    basliklar = [cell.value for cell in ws[1]]

    # Aynı başlığa sahip sütunların indekslerini belirle
    sutun_indexleri = {}
    for i, baslik in enumerate(basliklar):
        if baslik:
            # 'İTÜ Uzantılı E-posta adresiniz' başlığını 'E-posta adresiniz' ile birleştir
            anahtar = "E-posta adresiniz" if baslik in ["E-posta adresiniz",
                                                        "İTÜ Uzantılı E-posta adresiniz"] else baslik
            if anahtar not in sutun_indexleri:
                sutun_indexleri[anahtar] = []
            sutun_indexleri[anahtar].append(i)

    # Telefon numarası sütunlarını belirle
    telefon_sutunlari = ["Telefon Numaranız (5xxxxxxxxx)"]

    # Satırları işle
    veriler = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        satir_verisi = {}
        for baslik, index_listesi in sutun_indexleri.items():
            degerler = [row[i] for i in index_listesi if row[i] is not None]

            # Telefon numarası ise string formatına çevir
            if baslik in telefon_sutunlari and degerler:
                satir_verisi[baslik] = str(int(degerler[0]))  # Ondalık kısımdan kurtul
            else:
                satir_verisi[baslik] = degerler[0] if degerler else None
        satir_verisi.pop("Eklemek istediğiniz bir şey var mı?", None)
        veriler.append(satir_verisi)
    return veriler


def add_tokens_to_veriler(veriler, existing_tokens):
    for veri in veriler:
        token = veri.get(TOKEN, None)
        if not token:
            token = generate_unique_code(existing_tokens)
            veri[TOKEN] = token
        existing_tokens.add(token)
    return veriler


def create_davetliler_xlsx(veriler, dosya_adi):
    # Excel dosyasını oluştur
    wb = openpyxl.Workbook()
    ws = wb.active

    # Sütun başlıklarını belirle (dictionary'nin key'leri)
    sutun_basliklari = list(veriler[0].keys())
    ws.append(sutun_basliklari)

    # Verileri ekle
    for veri in veriler:
        ws.append([veri[key] for key in sutun_basliklari])

    # Yeni Excel dosyasını kaydet
    wb.save(dosya_adi)


def merge_yanitlar_ve_davetliler_verileri(veriler, mevcut_veriler):
    # İlk dosyadaki verileri ikinci dosyadaki verilerle birleştir, token'ları koru
    guncellenmis_veriler = []
    for veri in veriler:
        # Eğer bu veri zaten mevcut verilerde varsa, token'ı koru
        mevcut_veri = next((v for v in mevcut_veriler if v.get(ZAMAN_DAMGASI) == veri.get(ZAMAN_DAMGASI)), None)
        if mevcut_veri:
            veri[TOKEN] = mevcut_veri[TOKEN]
        guncellenmis_veriler.append(veri)
    return guncellenmis_veriler


FORM_YANITLARI_DOSYASI = "İTÜ İftarı - 2025 (Yanıtlar).xlsx"
DAVETLILER_DOSYASI = "itu_iftari_davetliler.xlsx"
TOKEN = "token"
OKUL = "Okulunuz"
ISIM = "Adınız Soyadınız"
TELEFON = "Telefon Numaranız (5xxxxxxxxx)"
ZAMAN_DAMGASI = "Zaman damgası"

# İlk dosyayı oku ve parse et
veriler = read_and_parse_form_yanitlari_xlsx(FORM_YANITLARI_DOSYASI)
mevcut_veriler = []
if os.path.exists(DAVETLILER_DOSYASI):
    mevcut_veriler = read_and_parse_form_yanitlari_xlsx(DAVETLILER_DOSYASI)

guncellenmis_veriler = merge_yanitlar_ve_davetliler_verileri(veriler, mevcut_veriler)
mevcut_tokenlar = set(veri[TOKEN] for veri in guncellenmis_veriler if TOKEN in veri)

# Yeni eklenen kayıtlara token atama işlemini gerçekleştir
guncellenmis_veriler = add_tokens_to_veriler(guncellenmis_veriler, mevcut_tokenlar)

# Güncellenmiş verilerle .js dosyasını ve ikinci Excel dosyasını oluştur
create_davetli_listesi_js(guncellenmis_veriler)
create_davetliler_xlsx(guncellenmis_veriler, DAVETLILER_DOSYASI)

# Sonuçları yazdır
for veri in guncellenmis_veriler[:10]:
    print(veri[TOKEN], "\t-\t", veri[TELEFON], "\t-\t", veri[OKUL], "\t-\t", veri[ISIM])
