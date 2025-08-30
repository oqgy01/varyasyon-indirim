import requests
import xml.etree.ElementTree as ET
import pandas as pd
import time
from typing import List, Dict, Any



def get_xml_data(url: str, max_retries: int = 10) -> str:
    """
    Belirtilen URL'den XML verisini indirir.
    Hata durumunda 5 saniye bekleyip tekrar dener.
    """
    for attempt in range(max_retries):
        try:
            print(f"İstek gönderiliyor: {url} (Deneme {attempt + 1})")
            response = requests.get(url, timeout=9999)
            response.raise_for_status()
            print(f"Başarılı: {url}")
            return response.text
        except Exception as e:
            print(f"Hata (Deneme {attempt + 1}): {url} - {str(e)}")
            if attempt < max_retries - 1:
                print("5 saniye bekleniyor...")
                time.sleep(5)
            else:
                print(f"Maksimum deneme sayısına ulaşıldı: {url}")
                raise e

def parse_xml_products(xml_content: str) -> List[Dict[str, Any]]:
    """
    XML içeriğini parse eder ve ürün verilerini liste olarak döner.
    """
    products = []
    try:
        root = ET.fromstring(xml_content)
        
        for product in root.findall('.//Product'):
            product_data = {}
            
            # Her bir XML alanını kontrol edip değeri al
            fields = [
                'IdUrun', 'UrunAdi', 'StokKodu', 'SatistakiStokAdedi',
                'SatistaOlduguGunlerVeBedenlerinSatistakiStokAdetleri',
                    'Kategori', 'Mevsim', 'UrununAktifBedenOrani', 'GuncelSatisFiyati'
            ]
            
            for field in fields:
                element = product.find(field)
                if element is not None:
                    product_data[field] = element.text.strip() if element.text else ""
                else:
                    product_data[field] = ""
            
            products.append(product_data)
            
    except ET.ParseError as e:
        print(f"XML parse hatası: {str(e)}")
        return []
    
    return products

def filter_products(products: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Ürünleri belirtilen kriterlere göre filtreler:
    1. SatistaOlduguGunlerVeBedenlerinSatistakiStokAdetleri kolonunda en az 2 adet // içerenler
    2. SatistakiStokAdedi 25'ten küçük olanlar silinir
    3. UrununAktifBedenOrani 51'den küçük olanlar silinir
    """
    filtered_products = []
    
    for product in products:
        # 1. SatistaOlduguGunlerVeBedenlerinSatistakiStokAdetleri kontrolü
        stok_bedenler = product.get('SatistaOlduguGunlerVeBedenlerinSatistakiStokAdetleri', '')
        if stok_bedenler.count('//') < 2:
            continue  # Bu ürünü atla
        
        # 2. SatistakiStokAdedi kontrolü
        try:
            satistaki_stok = int(product.get('SatistakiStokAdedi', '0'))
            if satistaki_stok < 25:
                continue  # Bu ürünü atla
        except (ValueError, TypeError):
            continue  # Sayıya çevrilemiyorsa atla
        
        # 3. UrununAktifBedenOrani kontrolü
        try:
            aktif_beden_orani = int(product.get('UrununAktifBedenOrani', '0'))
            if aktif_beden_orani < 51:
                continue  # Bu ürünü atla
        except (ValueError, TypeError):
            continue  # Sayıya çevrilemiyorsa atla
        
        # Tüm kriterleri geçen ürünü listeye ekle
        filtered_products.append(product)
    
    return filtered_products

def merge_excel_data():
    """
    urun_verileri.xlsx ve islenmis_veriler.xlsx dosyalarını birleştirir.
    StokKodu eşleşmesi yaparak beden stok bilgilerini günceller.
    """
    try:
        print("\nExcel dosyaları birleştiriliyor...")
        
        # Excel dosyalarını oku
        urun_df = pd.read_excel("urun_verileri.xlsx")
        islenmis_df = pd.read_excel("islenmis_veriler.xlsx")
        
        print(f"urun_verileri.xlsx: {len(urun_df)} satır")
        print(f"islenmis_veriler.xlsx: {len(islenmis_df)} satır")
        
        # StokKodu eşleşmesi yap
        def update_beden_stok(row):
            stok_kodu = row['StokKodu']
            beden_stok_str = row['SatistaOlduguGunlerVeBedenlerinSatistakiStokAdetleri']
            
            if pd.isna(beden_stok_str) or not isinstance(beden_stok_str, str):
                return beden_stok_str
            
            # Beden stok verilerini parçala
            beden_parts = beden_stok_str.split(' // ')
            updated_parts = []
            
            for part in beden_parts:
                if ' : ' in part:
                    beden, stok = part.split(' : ')
                    beden = beden.strip()
                    stok = stok.strip()
                    
                    # Bu bedeni islenmis_veriler.xlsx'de ara
                    matching_rows = islenmis_df[
                        (islenmis_df['StokKoduDuzenlenmis'] == stok_kodu) & 
                        (islenmis_df['Varyant'] == beden)
                    ]
                    
                    if not matching_rows.empty:
                        # Eşleşme bulundu, EtoplaAdet değerini al
                        etopla_adet = matching_rows.iloc[0]['EtoplaAdet']
                        updated_part = f"{beden} : {stok}-{int(etopla_adet)}"
                    else:
                        # Eşleşme bulunamadı, 0 ekle
                        updated_part = f"{beden} : {stok}-0"
                    
                    updated_parts.append(updated_part)
                else:
                    updated_parts.append(part)
            
            return ' // '.join(updated_parts)
        
        # Beden stok kolonunu güncelle
        urun_df['SatistaOlduguGunlerVeBedenlerinSatistakiStokAdetleri'] = urun_df.apply(update_beden_stok, axis=1)
        
        # Beden oranlarını hesapla
        print("\nBeden oranları hesaplanıyor...")
        urun_df = calculate_beden_ratios(urun_df)
        
        # SismeOrani kolonunu ekle
        print("\nSismeOrani kolonu ekleniyor...")
        urun_df = calculate_sisme_orani(urun_df)
        
        # SismeOrani 40'tan küçük değerleri filtrele
        print("\nSismeOrani 40'tan küçük değerler filtreleniyor...")
        urun_df = filter_sisme_orani(urun_df)
        
        # Supabase'e bağlan
        print("\nSupabase veritabanına bağlanılıyor...")
        supabase = connect_supabase()
        
        if supabase:
            # SatisaGirmeTarihi verilerini çek
            print("\nSatisaGirmeTarihi verileri çekiliyor...")
            urun_df = get_satisa_girme_tarihi(urun_df, supabase)
            
            # Son 5 gün içindeki tarihleri filtrele
            print("\nSon 5 gün içindeki tarihler filtreleniyor...")
            urun_df = filter_recent_dates(urun_df)
        
        # S ve 36 bedenlerini temizle
        print("\nS ve 36 bedenleri temizleniyor...")
        urun_df = clean_beden_names(urun_df)
        
        # VaryantFiyati kolonunu ekle
        print("\nVaryantFiyati kolonu ekleniyor...")
        urun_df = calculate_varyant_fiyati(urun_df)
        
        # Güncellenmiş dosyayı kaydet
        output_filename = "guncellenmis_urun_verileri.xlsx"
        urun_df.to_excel(output_filename, index=False, engine='openpyxl')
        
        print(f"\n✅ Excel dosyaları başarıyla birleştirildi!")
        print(f"📁 Güncellenmiş dosya: {output_filename}")
        print(f"📊 Toplam satır: {len(urun_df)}")
        
        return True
        
    except Exception as e:
        print(f"\n❌ Excel birleştirme hatası: {str(e)}")
        return False

def calculate_beden_ratios(df: pd.DataFrame) -> pd.DataFrame:
    """
    SatistaOlduguGunlerVeBedenlerinSatistakiStokAdetleri kolonundaki verileri hesaplar:
    L : 14-2 // M : 53-2 // S : 31-1 → L : 7 // M : 27 // S : 31
    Soldaki sayıyı sağdakine böler, sonucu en yakın tam sayıya yuvarlar.
    Sağdaki sayı 0 ise soldaki sayıyı olduğu gibi bırakır.
    """
    try:
        def calculate_ratio(beden_stok_str):
            if pd.isna(beden_stok_str) or not isinstance(beden_stok_str, str):
                return beden_stok_str
            
            # Beden stok verilerini parçala
            beden_parts = beden_stok_str.split(' // ')
            updated_parts = []
            
            for part in beden_parts:
                if ' : ' in part:
                    beden, stok = part.split(' : ')
                    beden = beden.strip()
                    stok = stok.strip()
                    
                    # Stok değerini parçala (örn: "14-2")
                    if '-' in stok:
                        try:
                            left_num, right_num = stok.split('-')
                            left_num = int(left_num.strip())
                            right_num = int(right_num.strip())
                            
                            # Sağdaki sayı 0 ise soldaki sayıyı olduğu gibi bırak
                            if right_num == 0:
                                result = left_num
                            else:
                                # Soldaki sayıyı sağdakine böl ve en yakın tam sayıya yuvarla
                                result = round(left_num / right_num)
                            
                            updated_part = f"{beden} : {result}"
                        except (ValueError, ZeroDivisionError):
                            # Hata durumunda orijinal değeri koru
                            updated_part = part
                    else:
                        # "-" yoksa orijinal değeri koru
                        updated_part = part
                    
                    updated_parts.append(updated_part)
                else:
                    updated_parts.append(part)
            
            return ' // '.join(updated_parts)
        
        # Beden stok kolonunu güncelle
        df['SatistaOlduguGunlerVeBedenlerinSatistakiStokAdetleri'] = df['SatistaOlduguGunlerVeBedenlerinSatistakiStokAdetleri'].apply(calculate_ratio)
        
        print("✅ Beden oranları hesaplandı ve güncellendi!")
        return df
        
    except Exception as e:
        print(f"❌ Beden oranları hesaplama hatası: {str(e)}")
        return df

def calculate_sisme_orani(df: pd.DataFrame) -> pd.DataFrame:
    """
    SismeOrani kolonunu ekler ve 36/S bedeninin diğer bedenlere olan ortalama uzaklık yüzdesini hesaplar.
    Sadece stok değeri en az 10 olan bedenler karşılaştırmaya dahil edilir.
    Örnek: 36:17 // 38:17 // 40:18 // 42:16 // 44:18 // 46:16
    - 36(17) ile 38(17): %0 (17-17)/17 * 100 = 0% (17>=10 ✓)
    - 36(17) ile 40(18): %-5.88 (17-18)/17 * 100 = -5.88% (18>=10 ✓)
    - 36(17) ile 42(16): %5.88 (17-16)/17 * 100 = 5.88% (16>=10 ✓)
    - vs... sonra ortalaması alınır
    """
    try:
        def calculate_sisme_percentage(beden_stok_str):
            if pd.isna(beden_stok_str) or not isinstance(beden_stok_str, str):
                return None
            
            # Beden stok verilerini parçala
            beden_parts = beden_stok_str.split(' // ')
            percentages = []
            
            # Referans beden (36 veya S) bul
            reference_beden = None
            reference_value = None
            
            for part in beden_parts:
                if ' : ' in part:
                    beden, stok = part.split(' : ')
                    beden = beden.strip()
                    stok = stok.strip()
                    
                    # Referans beden kontrolü (36 veya S)
                    if beden == '36' or beden == 'S':
                        try:
                            reference_beden = beden
                            reference_value = int(stok)
                            break
                        except ValueError:
                            continue
            
            # Referans beden bulunamadıysa None döndür
            if reference_beden is None or reference_value is None:
                return None
            
            # Diğer bedenlerle karşılaştır ve yüzde hesapla
            for part in beden_parts:
                if ' : ' in part:
                    beden, stok = part.split(' : ')
                    beden = beden.strip()
                    stok = stok.strip()
                    
                    # Referans beden değilse hesapla
                    if beden != reference_beden:
                        try:
                            compare_value = int(stok)
                            
                            # Sadece stok değeri en az 10 olan bedenleri hesaplamaya dahil et
                            if compare_value >= 10:
                                # Yüzde hesapla: (referans - karşılaştırılan) / referans * 100
                                if reference_value != 0:  # Sıfıra bölme kontrolü
                                    percentage = ((reference_value - compare_value) / reference_value) * 100
                                    percentages.append(percentage)
                        except ValueError:
                            continue
            
            # Ortalama yüzdeyi hesapla
            if percentages:
                average_percentage = sum(percentages) / len(percentages)
                return round(average_percentage, 2)  # 2 ondalık basamağa yuvarla
            else:
                return None
        
        # Yeni kolonu ekle
        df['SismeOrani'] = df['SatistaOlduguGunlerVeBedenlerinSatistakiStokAdetleri'].apply(calculate_sisme_percentage)
        
        print("✅ SismeOrani kolonu başarıyla eklendi!")
        print(f"📊 Toplam {len(df)} satırdan {df['SismeOrani'].notna().sum()} satırda SismeOrani hesaplandı")
        return df
        
    except Exception as e:
        print(f"❌ SismeOrani hesaplama hatası: {str(e)}")
        return df

def filter_sisme_orani(df: pd.DataFrame) -> pd.DataFrame:
    """
    SismeOrani kolonunda 40'tan küçük değerleri olan satırları siler.
    """
    try:
        initial_rows = len(df)
        
        # SismeOrani kolonunda 40'tan küçük değerleri olan satırları sil
        df = df[df['SismeOrani'] >= 40]
        
        final_rows = len(df)
        removed_rows = initial_rows - final_rows
        
        print(f"✅ SismeOrani filtrelendi!")
        print(f"📊 {removed_rows} satır silindi (40'tan küçük değerler)")
        print(f"📊 Kalan satır: {final_rows}")
        
        return df
        
    except Exception as e:
        print(f"❌ SismeOrani filtreleme hatası: {str(e)}")
        return df

def connect_supabase():
    """
    Supabase veritabanına bağlanır.
    """
    try:
        from supabase import create_client, Client
        
        # Supabase bağlantı bilgileri
        SUPABASE_URL = "https://zmvsatlvobhdaxxgtoap.supabase.co"
        SUPABASE_KEY = (
            "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9."
            "eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InptdnNhdGx2b2JoZGF4eGd0b2FwIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NDAxNzIxMzksImV4cCI6MjA1NTc0ODEzOX0."
            "lJLudSfixMbEOkJmfv22MsRLofP7ZjFkbGj26xF3dts"
        )
        
        # Supabase istemcisini oluştur
        supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
        
        print("✅ Supabase veritabanına başarıyla bağlandı!")
        return supabase
        
    except ImportError:
        print("❌ Supabase kütüphanesi bulunamadı! 'pip install supabase' komutunu çalıştırın.")
        return None
    except Exception as e:
        print(f"❌ Supabase bağlantı hatası: {str(e)}")
        return None

def get_satisa_girme_tarihi(df: pd.DataFrame, supabase) -> pd.DataFrame:
    """
    Supabase'den SatisaGirmeTarihi verilerini çeker ve yeni kolon olarak ekler.
    """
    try:
        if supabase is None:
            print("❌ Supabase bağlantısı bulunamadı!")
            return df
        
        print("📊 SatisaGirmeTarihi verileri çekiliyor...")
        
        # Yeni kolonu ekle
        df['SatisaGirmeTarihi'] = None
        
        # Her StokKodu için veritabanında ara
        for index, row in df.iterrows():
            stok_kodu = row['StokKodu']
            
            try:
                # "indirim-bindirim" tablosunda StokKodu ara
                response = supabase.table("indirim-bindirim").select("SatisaGirmeTarihi").eq("StokKodu", stok_kodu).execute()
                
                if response.data and len(response.data) > 0:
                    # İlk eşleşen kaydın SatisaGirmeTarihi'ni al
                    satisa_girme_tarihi = response.data[0]['SatisaGirmeTarihi']
                    df.at[index, 'SatisaGirmeTarihi'] = satisa_girme_tarihi
                
            except Exception as e:
                print(f"⚠️ StokKodu {stok_kodu} için veri çekilemedi: {str(e)}")
                continue
        
        # Başarılı şekilde veri çekilen satır sayısını göster
        successful_rows = df['SatisaGirmeTarihi'].notna().sum()
        print(f"✅ SatisaGirmeTarihi kolonu eklendi!")
        print(f"📊 {successful_rows} satırda veri bulundu")
        
        return df
        
    except Exception as e:
        print(f"❌ SatisaGirmeTarihi çekme hatası: {str(e)}")
        return df

def filter_recent_dates(df: pd.DataFrame) -> pd.DataFrame:
    """
    SatisaGirmeTarihi kolonunda son 5 gün içindeki tarihleri olan satırları siler.
    """
    try:
        from datetime import datetime, timedelta
        
        initial_rows = len(df)
        
        # Bugünün tarihini al
        today = datetime.now().date()
        
        # Son 5 günü hesapla
        five_days_ago = today - timedelta(days=5)
        
        # SatisaGirmeTarihi kolonundaki tarihleri kontrol et
        rows_to_remove = []
        
        for index, row in df.iterrows():
            satisa_girme_tarihi = row['SatisaGirmeTarihi']
            
            if pd.notna(satisa_girme_tarihi) and isinstance(satisa_girme_tarihi, str):
                try:
                    # String tarihi datetime objesine çevir
                    if 'T' in satisa_girme_tarihi:  # ISO format
                        tarih = datetime.fromisoformat(satisa_girme_tarihi.replace('Z', '+00:00')).date()
                    else:  # Diğer formatlar
                        tarih = datetime.strptime(satisa_girme_tarihi, '%Y-%m-%d').date()
                    
                    # Son 5 gün içindeyse silinecek satırlara ekle
                    if tarih >= five_days_ago:
                        rows_to_remove.append(index)
                        
                except (ValueError, TypeError):
                    # Tarih parse edilemiyorsa satırı koru
                    continue
        
        # Son 5 gün içindeki satırları sil
        if rows_to_remove:
            df = df.drop(rows_to_remove)
            df = df.reset_index(drop=True)
        
        final_rows = len(df)
        removed_rows = initial_rows - final_rows
        
        print(f"✅ Son 5 gün içindeki tarihler filtrelendi!")
        print(f"📊 {removed_rows} satır silindi (son 5 gün içindeki tarihler)")
        print(f"📊 Kalan satır: {final_rows}")
        
        return df
        
    except Exception as e:
        print(f"❌ Tarih filtreleme hatası: {str(e)}")
        return df

def clean_beden_names(df: pd.DataFrame) -> pd.DataFrame:
    """
    SatistaOlduguGunlerVeBedenlerinSatistakiStokAdetleri kolonunda S ve 36 bedenlerini temizler.
    Sadece beden adını bırakır.
    """
    try:
        def clean_beden_stok_str(beden_stok_str):
            if pd.isna(beden_stok_str) or not isinstance(beden_stok_str, str):
                return beden_stok_str
            
            # Beden stok verilerini parçala
            beden_parts = beden_stok_str.split(' // ')
            updated_parts = []
            
            for part in beden_parts:
                if ' : ' in part:
                    beden, stok = part.split(' : ')
                    beden = beden.strip()
                    stok = stok.strip()
                    
                    # S veya 36 bedenlerini sadece beden adı olarak bırak
                    if beden == 'S' or beden == '36':
                        updated_part = beden
                    else:
                        # Diğer bedenler için orijinal formatı koru
                        updated_part = f"{beden} : {stok}"
                    
                    updated_parts.append(updated_part)
                else:
                    updated_parts.append(part)
            
            return ' // '.join(updated_parts)
        
        # Beden stok kolonunu güncelle
        df['SatistaOlduguGunlerVeBedenlerinSatistakiStokAdetleri'] = df['SatistaOlduguGunlerVeBedenlerinSatistakiStokAdetleri'].apply(clean_beden_stok_str)
        
        print("✅ S ve 36 bedenleri temizlendi!")
        return df
        
    except Exception as e:
        print(f"❌ Beden temizleme hatası: {str(e)}")
        return df

def calculate_varyant_fiyati(df: pd.DataFrame) -> pd.DataFrame:
    """
    VaryantFiyati kolonunu ekler ve SismeOrani'na göre fiyat hesaplaması yapar.
    SismeOrani 40-70 arası: %15 indirim
    SismeOrani 70+ : %20 indirim
    Sonuç yuvarlama kodu ile yuvarlanır.
    """
    try:
        def round_price(price):
            """
            Fiyat yuvarlama kodu - JavaScript'ten Python'a çevrildi
            """
            if pd.isna(price) or not isinstance(price, (int, float)) or price <= 0:
                return price
            
            # Önce belirli aralıkları kontrol et
            if 100 <= price <= 105:
                return 99.99
            elif 200 <= price <= 207:
                return 199.99
            elif 300 <= price <= 309:
                return 299.99
            elif 400 <= price <= 412:
                return 399.99
            elif 500 <= price <= 520:
                return 499.99
            
            # Aralık dışında ise normal yuvarlama işlemi
            base = int(price)
            price_tens = (base // 10) * 10
            target1 = price_tens - 5.01
            target2 = price_tens - 0.01
            target3 = price_tens + 4.99
            target4 = price_tens + 9.99
            target5 = price_tens + 14.99
            
            # Pozitif hedefleri filtrele
            targets = [t for t in [target1, target2, target3, target4, target5] if t > 0]
            
            if not targets:
                return None
            
            # En yakın hedefi bul
            closest_target = targets[0]
            min_difference = abs(price - closest_target)
            
            for target in targets[1:]:
                difference = abs(price - target)
                if difference < min_difference:
                    min_difference = difference
                    closest_target = target
            
            return closest_target
        
        def calculate_discounted_price(row):
            sisme_orani = row['SismeOrani']
            guncel_fiyat = row['GuncelSatisFiyati']
            
            if pd.isna(sisme_orani) or pd.isna(guncel_fiyat):
                return None
            
            try:
                # Fiyatı sayıya çevir
                if isinstance(guncel_fiyat, str):
                    # Virgülü nokta ile değiştir
                    guncel_fiyat = guncel_fiyat.replace(',', '.')
                
                fiyat = float(guncel_fiyat)
                
                # SismeOrani'na göre indirim uygula
                if 40 <= sisme_orani <= 70:
                    # %15 indirim
                    indirimli_fiyat = fiyat * 0.85
                elif sisme_orani > 70:
                    # %20 indirim
                    indirimli_fiyat = fiyat * 0.80
                else:
                    # İndirim yok
                    indirimli_fiyat = fiyat
                
                # Yuvarlama kodu ile yuvarla
                final_price = round_price(indirimli_fiyat)
                return final_price
                
            except (ValueError, TypeError):
                return None
        
        # Yeni kolonu ekle
        df['VaryantFiyati'] = df.apply(calculate_discounted_price, axis=1)
        
        # İstatistikleri göster
        total_rows = len(df)
        calculated_rows = df['VaryantFiyati'].notna().sum()
        
        print("✅ VaryantFiyati kolonu başarıyla eklendi!")
        print(f"�� Toplam {total_rows} satırdan {calculated_rows} satırda fiyat hesaplandı")
        
        return df
        
    except Exception as e:
        print(f"❌ VaryantFiyati hesaplama hatası: {str(e)}")
        return df


def main():
    # İşlenecek linkler
    urls = [
        "https://www.siparis.haydigiy.com/FaprikaXml/2XO5DS/1/",
        "https://www.siparis.haydigiy.com/FaprikaXml/2XO5DS/2/",
        "https://www.siparis.haydigiy.com/FaprikaXml/2XO5DS/3/"
    ]
    
    all_products = []
    
    print("XML verileri indiriliyor...")
    print("=" * 50)
    
    # Her linki sırayla işle
    for i, url in enumerate(urls, 1):
        print(f"\n{i}. Link işleniyor...")
        try:
            xml_content = get_xml_data(url)
            products = parse_xml_products(xml_content)
            all_products.extend(products)
            print(f"{len(products)} ürün bulundu")
        except Exception as e:
            print(f"Link işlenemedi: {url} - Hata: {str(e)}")
            continue
    
    print("\n" + "=" * 50)
    print(f"Toplam {len(all_products)} ürün verisi toplandı")
    
    # Filtreleme işlemi
    print("\nÜrünler filtreleniyor...")
    filtered_products = filter_products(all_products)
    print(f"Filtreleme sonrası {len(filtered_products)} ürün kaldı")
    
    if filtered_products:
        # DataFrame oluştur
        df = pd.DataFrame(filtered_products)
        
        # Excel dosyasına kaydet
        excel_filename = "urun_verileri.xlsx"
        df.to_excel(excel_filename, index=False, engine='openpyxl')
        print(f"\nVeriler başarıyla '{excel_filename}' dosyasına kaydedildi!")
        

        
    else:
        print("Hiç ürün verisi bulunamadı!")

if __name__ == "__main__":
    main()























# İKİNCİ KISIM

import requests
import pandas as pd
import time
from typing import Dict, Any

def download_excel_file(url: str, max_retries: int = 10) -> bytes:
    """
    Belirtilen URL'den Excel dosyasını indirir.
    Hata durumunda 5 saniye bekleyip tekrar dener.
    """
    for attempt in range(max_retries):
        try:
            print(f"Excel dosyası indiriliyor: {url} (Deneme {attempt + 1})")
            response = requests.get(url, timeout=9999)
            response.raise_for_status()
            print(f"Excel dosyası başarıyla indirildi!")
            return response.content
        except Exception as e:
            print(f"Hata (Deneme {attempt + 1}): {url} - {str(e)}")
            if attempt < max_retries - 1:
                print("5 saniye bekleniyor...")
                time.sleep(5)
            else:
                print(f"Maksimum deneme sayısına ulaşıldı: {url}")
                raise e

def process_excel_data(excel_content: bytes) -> pd.DataFrame:
    """
    Excel içeriğini işler ve gerekli kolonları filtreler.
    """
    try:
        # Excel dosyasını oku
        df = pd.read_excel(excel_content, engine='openpyxl')
        print(f"Excel dosyası okundu. Toplam {len(df)} satır ve {len(df.columns)} kolon bulundu.")
        
        # Sadece gerekli kolonları tut
        required_columns = ['StokKodu', 'Adet', 'Varyant']
        
        # Eksik kolonları kontrol et
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            print(f"Uyarı: Eksik kolonlar: {missing_columns}")
            return pd.DataFrame()
        
        # Sadece gerekli kolonları seç
        df_filtered = df[required_columns].copy()
        print(f"Filtreleme sonrası {len(df_filtered)} satır kaldı.")
        
        return df_filtered
        
    except Exception as e:
        print(f"Excel işleme hatası: {str(e)}")
        return pd.DataFrame()

def add_etopla_adet_column(df: pd.DataFrame) -> pd.DataFrame:
    """
    EtoplaAdet kolonunu ekler - StokKodu'na göre gruplayıp Adet'leri toplar.
    """
    try:
        # Adet kolonundaki verileri sayıya çevir
        def convert_adet_to_number(adet_value):
            if pd.isna(adet_value):
                return 0
            try:
                # String ise, virgülü nokta ile değiştir ve float'a çevir
                if isinstance(adet_value, str):
                    # Virgülü nokta ile değiştir
                    adet_value = adet_value.replace(',', '.')
                    return float(adet_value)
                else:
                    return float(adet_value)
            except (ValueError, TypeError):
                return 0
        
        # Adet kolonunu sayıya çevir
        df['Adet_Numeric'] = df['Adet'].apply(convert_adet_to_number)
        
        # StokKodu'na göre grupla ve Adet'leri topla
        etopla_dict = df.groupby('StokKodu')['Adet_Numeric'].sum().to_dict()
        
        # Yeni kolonu ekle
        df['EtoplaAdet'] = df['StokKodu'].map(etopla_dict)
        
        # Geçici kolonu sil
        df = df.drop('Adet_Numeric', axis=1)
        
        print("EtoplaAdet kolonu başarıyla eklendi.")
        return df
        
    except Exception as e:
        print(f"EtoplaAdet kolonu ekleme hatası: {str(e)}")
        return df

def add_stok_kodu_duzenlenmis_column(df: pd.DataFrame) -> pd.DataFrame:
    """
    StokKoduDuzenlenmis kolonunu ekler - 3. noktadan sonrasını temizler.
    """
    try:
        def clean_stok_kodu(stok_kodu: str) -> str:
            """StokKodu'ndan 3. noktadan sonrasını temizler."""
            if pd.isna(stok_kodu) or not isinstance(stok_kodu, str):
                return stok_kodu
            
            # Nokta sayısını say
            dot_count = stok_kodu.count('.')
            
            if dot_count >= 3:
                # 3. noktaya kadar olan kısmı al
                parts = stok_kodu.split('.')
                return '.'.join(parts[:3])
            else:
                # 3 noktadan az ise olduğu gibi bırak
                return stok_kodu
        
        # Yeni kolonu ekle
        df['StokKoduDuzenlenmis'] = df['StokKodu'].apply(clean_stok_kodu)
        
        print("StokKoduDuzenlenmis kolonu başarıyla eklendi.")
        return df
        
    except Exception as e:
        print(f"StokKoduDuzenlenmis kolonu ekleme hatası: {str(e)}")
        return df

def clean_varyant_column(df: pd.DataFrame) -> pd.DataFrame:
    """
    Varyant kolonundaki "Beden: " kısmını temizler.
    """
    try:
        def clean_varyant(varyant_value):
            if pd.isna(varyant_value) or not isinstance(varyant_value, str):
                return varyant_value
            
            # "Beden: " kısmını kaldır
            if varyant_value.startswith("Beden: "):
                return varyant_value.replace("Beden: ", "")
            else:
                return varyant_value
        
        # Varyant kolonunu temizle
        df['Varyant'] = df['Varyant'].apply(clean_varyant)
        
        print("Varyant kolonu temizlendi.")
        return df
        
    except Exception as e:
        print(f"Varyant kolonu temizleme hatası: {str(e)}")
        return df

def remove_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    StokKodu ve Adet kolonlarını siler.
    """
    try:
        # Silinecek kolonlar
        columns_to_remove = ['StokKodu', 'Adet']
        
        # Mevcut kolonları kontrol et ve sil
        existing_columns = [col for col in columns_to_remove if col in df.columns]
        if existing_columns:
            df = df.drop(columns=existing_columns, axis=1)
            print(f"Kolonlar silindi: {existing_columns}")
        else:
            print("Silinecek kolon bulunamadı.")
        
        return df
        
    except Exception as e:
        print(f"Kolon silme hatası: {str(e)}")
        return df

def remove_duplicates(df: pd.DataFrame) -> pd.DataFrame:
    """
    Tekrarlanan satırları kaldırır.
    """
    try:
        # Tekrarlanan satırları kaldır
        initial_rows = len(df)
        df = df.drop_duplicates()
        final_rows = len(df)
        removed_rows = initial_rows - final_rows
        
        if removed_rows > 0:
            print(f"Tekrarlanan {removed_rows} satır kaldırıldı.")
        else:
            print("Tekrarlanan satır bulunamadı.")
        
        return df
        
    except Exception as e:
        print(f"Tekrarlanan satır kaldırma hatası: {str(e)}")
        return df

def main():
    # İndirilecek Excel dosyası URL'i
    url = "https://www.siparis.haydigiy.com/FaprikaOrderXls/T6PPZN/1/"
    
    print("Excel işleme programı başlatılıyor...")
    print("=" * 60)
    
    try:
        # 1. Excel dosyasını indir
        print("\n1. Excel dosyası indiriliyor...")
        excel_content = download_excel_file(url)
        
        # 2. Excel verilerini işle ve filtrele
        print("\n2. Excel verileri işleniyor...")
        df = process_excel_data(excel_content)
        
        if df.empty:
            print("Excel verisi işlenemedi!")
            return
        
        # 3. EtoplaAdet kolonunu ekle
        print("\n3. EtoplaAdet kolonu ekleniyor...")
        df = add_etopla_adet_column(df)
        
        # 4. StokKoduDuzenlenmis kolonunu ekle
        print("\n4. StokKoduDuzenlenmis kolonu ekleniyor...")
        df = add_stok_kodu_duzenlenmis_column(df)
        
        # 5. Varyant kolonunu temizle
        print("\n5. Varyant kolonu temizleniyor...")
        df = clean_varyant_column(df)
        
        # 6. Gereksiz kolonları sil
        print("\n6. Gereksiz kolonlar siliniyor...")
        df = remove_columns(df)
        
        # 7. Tekrarlanan satırları kaldır
        print("\n7. Tekrarlanan satırlar kaldırılıyor...")
        df = remove_duplicates(df)
        
        # 8. Sonucu Excel olarak kaydet
        print("\n8. Sonucu Excel olarak kaydediliyor...")
        output_filename = "islenmis_veriler.xlsx"
        df.to_excel(output_filename, index=False, engine='openpyxl')
        
        # 9. Excel dosyalarını birleştir
        print("\n9. Excel dosyaları birleştiriliyor...")
        merge_excel_data()
        
    except Exception as e:
        print(f"\n❌ Program hatası: {str(e)}")

if __name__ == "__main__":
    main()


















# SELENİUM İŞLEMLER


import requests
import time
import xml.etree.ElementTree as ET
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import re
import pandas as pd

# ─────────── GİRİŞ BİLGİLERİ ───────────
USER   = "mustafa_kod@haydigiy.com"
PASSWD = "123456"
# ────────────────────────────────────────

# ─────────── URL'LER ───────────
BASE_URL     = "https://www.siparis.haydigiy.com"
LOGIN_URL    = f"{BASE_URL}/kullanici-giris/?ReturnUrl=%2Fadmin"
BULKEDIT_URL = f"{BASE_URL}/admin/product/bulkedit/"
XML_URL = "https://www.siparis.haydigiy.com/FaprikaXml/NE6ZAB/1/"
# ───────────────────────────────

def get_xml_data():
    """XML verisini alır ve ürün ID'lerini döndürür."""
    max_retries = 10
    retry_count = 0
    
    while retry_count < max_retries:
        try:
            print(f"XML verisi alınıyor... (Deneme {retry_count + 1})")
            response = requests.get(XML_URL, timeout=9999)
            
            if response.status_code == 200:
                print("XML verisi başarıyla alındı!")
                xml_content = response.text
                
                # XML'i parse et
                root = ET.fromstring(xml_content)
                
                # Tüm ürün ID'lerini bul
                product_ids = []
                
                # Namespace ile birlikte arama yap
                for item in root.findall('.//item'):
                    # Önce namespace ile dene
                    product_id = item.find('{http://base.google.com/ns/1.0}id')
                    if product_id is not None and product_id.text:
                        product_ids.append(product_id.text)
                        continue
                    
                    # Namespace olmadan da dene
                    product_id = item.find('g:id')
                    if product_id is not None and product_id.text:
                        product_ids.append(product_id.text)
                        continue
                    
                    # Direkt id olarak da dene
                    product_id = item.find('id')
                    if product_id is not None and product_id.text:
                        product_ids.append(product_id.text)
                        continue
                
                # Debug için XML içeriğini yazdır
                print("XML içeriği:")
                print(xml_content[:500] + "..." if len(xml_content) > 500 else xml_content)
                
                print(f"Bulunan item sayısı: {len(root.findall('.//item'))}")
                for item in root.findall('.//item'):
                    print(f"Item içeriği: {ET.tostring(item, encoding='unicode')[:200]}...")
                
                print(f"Toplam {len(product_ids)} ürün ID'si bulundu.")
                return product_ids
                
            else:
                print(f"HTTP Hatası: {response.status_code}")
                retry_count += 1
                
        except Exception as e:
            print(f"XML alma hatası: {e}")
            retry_count += 1
            
        if retry_count < max_retries:
            print("5 saniye bekleniyor...")
            time.sleep(5)
    
    print("Maksimum deneme sayısına ulaşıldı. XML verisi alınamadı.")
    return []

def init_driver():
    """Tarayıcıyı (WebDriver) başlatır ve ayarlarını yapar."""
    opts = Options()
    opts.add_argument("--headless=new")  # Headless mod aktif
    opts.add_argument("--disable-gpu")
    opts.add_argument("--incognito")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument("--no-sandbox")  # Headless için ek güvenlik
    opts.add_argument("--disable-dev-shm-usage")  # Headless için ek güvenlik
    opts.add_experimental_option("excludeSwitches", ["enable-logging"])
    
    # Windows için Chrome yolu
    try:
        driver = webdriver.Chrome(options=opts)
        print("Chrome WebDriver başlatıldı.")
        return driver
    except Exception as e:
        print(f"Chrome WebDriver başlatılamadı: {e}")
        return None

def login(drv):
    """Admin paneline giriş yapar."""
    try:
        print("Giriş sayfasına gidiliyor...")
        drv.get(LOGIN_URL)
        
        print("E-posta/telefon alanı dolduruluyor...")
        email_field = WebDriverWait(drv, 15).until(
            EC.visibility_of_element_located((By.NAME, "EmailOrPhone"))
        )
        email_field.clear()
        email_field.send_keys(USER)
        
        print("Şifre alanı dolduruluyor...")
        password_field = drv.find_element(By.NAME, "Password")
        password_field.clear()
        password_field.send_keys(PASSWD)
        
        print("Giriş butonuna tıklanıyor...")
        login_button = drv.find_element(By.CSS_SELECTOR, "button[type='submit']")
        login_button.click()
        
        # Admin sayfasına yönlendirildiğini kontrol et
        WebDriverWait(drv, 15).until(EC.url_contains("/admin"))
        print("Giriş başarıyla yapıldı!")
        return True
        
    except Exception as e:
        print(f"Giriş hatası: {e}")
        return False

def bulk_edit_final_operations(drv):
    """Bulk edit sayfasında son işlemleri yapar."""
    try:
        print("\n=== BULK EDIT SON İŞLEMLERİ BAŞLIYOR ===")
        
        # Bulk edit sayfasına git
        print("Bulk edit sayfasına gidiliyor...")
        drv.get(BULKEDIT_URL)
        
        # Sayfa yüklenmesini bekle
        WebDriverWait(drv, 15).until(
            EC.presence_of_element_located((By.ID, "SearchInCategoryIds"))
        )
        
        # Kategori seçimi
        print("Kategori seçiliyor...")
        sel = Select(drv.find_element(By.ID, "SearchInCategoryIds"))
        sel.select_by_value("632")
        time.sleep(2)
        print("4 saniye beklendi - Kategori seçimi tamamlandı")
        
        # Fazla kategori seçimlerini temizle
        buttons = drv.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
        if len(buttons) > 1:
            buttons[1].click()
        time.sleep(2)
        print("4 saniye beklendi - Fazla kategoriler temizlendi")
        
        # Arama butonuna tıkla
        print("Ürün arama yapılıyor...")
        drv.find_element(By.ID, "search-products").click()
        time.sleep(2)
        print("4 saniye beklendi - Arama butonu tıklandı")
        
        # Ürün listesi yüklenmesini bekle
        WebDriverWait(drv, 15).until(
            EC.presence_of_element_located((By.ID, "ProductTag_Update"))
        )
        time.sleep(2)
        print("4 saniye beklendi - Ürün listesi yüklendi")
        
        # 1. ÜRÜN ETİKETİ İŞLEMLERİ
        print("Ürün etiketi işlemleri yapılıyor...")
        
        # ProductTag_Update checkbox'ını direkt click ile işaretle
        print("ProductTag_Update checkbox işaretleniyor...")
        chk = WebDriverWait(drv, 10).until(
            EC.presence_of_element_located((By.ID, "ProductTag_Update")))
        drv.execute_script("arguments[0].click();", chk)
        time.sleep(2)
        print("4 saniye beklendi - ProductTag_Update checkbox işaretlendi")
        
        # ProductTagId select2'den 241 ID'li değeri seç
        print("Etiket ID 241 seçiliyor...")
        drv.execute_script("""
            var $select = $("#ProductTagId");
            $select.val('241').trigger('change');
            $select.trigger('select2:select');
        """)
        time.sleep(2)
        print("4 saniye beklendi - Etiket ID 241 seçildi")
        
        # ProductTagTransactionId select2'den "Etiketi Çıkar" seç
        print("'Etiketi Çıkar' seçiliyor...")
        product_transaction_select = drv.find_element(By.ID, "ProductTagTransactionId")
        product_transaction_select = Select(product_transaction_select)
        product_transaction_select.select_by_value("1")
        time.sleep(2)
        print("4 saniye beklendi - 'Etiketi Çıkar' seçildi")
        
        # 2. KATEGORİ İŞLEMLERİ
        print("Kategori işlemleri yapılıyor...")
        
        # Category_Update checkbox'ını direkt click ile işaretle
        print("Category_Update checkbox işaretleniyor...")
        chk = WebDriverWait(drv, 10).until(
            EC.presence_of_element_located((By.ID, "Category_Update")))
        drv.execute_script("arguments[0].click();", chk)
        time.sleep(2)
        print("4 saniye beklendi - Category_Update checkbox işaretlendi")
        
        # CategoryId select2'den 632 ID'li değeri seç
        print("Kategori ID 632 seçiliyor...")
        drv.execute_script("""
            var $select = $("#CategoryId");
            $select.val('632').trigger('change');
            $select.trigger('select2:select');
        """)
        time.sleep(2)
        print("4 saniye beklendi - Kategori ID 632 seçildi")
        
        # CategoryTransactionId select2'den "Kategoriden Çıkar" seç
        print("'Kategoriden Çıkar' seçiliyor...")
        category_transaction_select = drv.find_element(By.ID, "CategoryTransactionId")
        category_transaction_select = Select(category_transaction_select)
        category_transaction_select.select_by_value("1")
        time.sleep(2)
        print("4 saniye beklendi - 'Kategoriden Çıkar' seçildi")
        
        # 30 saniye bekle
        print("4saniye bekleniyor...")
        time.sleep(2)
        
        # Sayfanın en üstüne çık
        print("Sayfanın en üstüne çıkılıyor...")
        drv.execute_script("window.scrollTo(0, 0);")
        time.sleep(1)
        
        # Kaydet butonuna tıkla
        print("Kaydet butonuna tıklanıyor...")
        save_button = WebDriverWait(drv, 15).until(
            EC.element_to_be_clickable((By.ID, "bulk-update-submit"))
        )
        save_button.click()
        
        print("Bulk edit işlemleri başarıyla tamamlandı!")
        return True
        
    except Exception as e:
        print(f"Bulk edit işlemlerinde hata: {e}")
        return False

def update_combination_prices_from_excel(drv):
    """Excel'den veri okuyup kombinasyon fiyatlarını günceller."""
    try:
        print("\n=== EXCEL'DEN KOMBİNASYON FİYATI GÜNCELLEME BAŞLIYOR ===")
        
        # Excel dosyasını oku
        print("Excel dosyası okunuyor...")
        df = pd.read_excel("guncellenmis_urun_verileri.xlsx")
        print(f"Excel'den {len(df)} satır okundu.")
        
        # Gerekli kolonları kontrol et
        required_columns = ['IdUrun', 'VaryantFiyati']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            print(f"Eksik kolonlar: {missing_columns}")
            return False
        
        successful_count = 0
        total_count = len(df)
        
        # Her satırı işle
        for index, row in df.iterrows():
            try:
                product_id = str(row['IdUrun'])
                variant_price = str(row['VaryantFiyati']).strip()
                
                print(f"\n--- Satır {index + 1}/{total_count} ---")
                print(f"Ürün ID: {product_id}")
                print(f"Varyant Fiyatı: {variant_price}")
                
                # Ürün düzenleme sayfasına git
                edit_url = f"{BASE_URL}/admin/product/edit/{product_id}"
                print(f"Ürün sayfasına gidiliyor: {edit_url}")
                drv.get(edit_url)
                
                # Sayfa yüklenmesini bekle
                WebDriverWait(drv, 15).until(
                    EC.presence_of_element_located((By.TAG_NAME, "body"))
                )
                
                # Önce ürün etiketi ekle (241 - S Bedeni İndirimli Ürünler)
                print("Ürün etiketi ekleniyor...")
                try:
                    # Select2 dropdown'ı bul ve 241 değerini seç
                    drv.execute_script("""
                        var $select = $("#SelectedProductTagIds");
                        if ($select.length > 0) {
                            $select.val('241').trigger('change');
                            $select.trigger('select2:select');
                        }
                    """)
                    print("Etiket 241 (S Bedeni İndirimli Ürünler) seçildi.")
                    
                    # Sayfanın en üstüne çık
                    drv.execute_script("window.scrollTo(0, 0);")

                    
                    # "Kaydet ve Devam Et" butonuna tıkla
                    save_continue_button = WebDriverWait(drv, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "button[name='save-continue']"))
                    )
                    save_continue_button.click()
                    print("Kaydet ve Devam Et butonuna tıklandı.")
                    
                    # Sayfa yeniden yüklenmesini bekle

                    print("5 saniye beklendi - Sayfa yeniden yüklendi.")
                    
                except Exception as e:
                    print(f"Ürün etiketi ekleme hatası: {e}")
                    continue
                
                # "Kategori / Marka" sekmesine tıkla
                print("Kategori / Marka sekmesi aranıyor...")
                try:
                    category_tab = WebDriverWait(drv, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//li[@data-tab-name='tab-mappings']"))
                    )
                    category_tab.click()
                    print("Kategori / Marka sekmesi tıklandı.")
                    
                    # "Yeni Kayıt Ekle" butonuna tıkla
                    print("Yeni Kayıt Ekle butonu aranıyor...")
                    add_button = WebDriverWait(drv, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "a.k-button.k-button-icontext.k-grid-add"))
                    )
                    add_button.click()
                    print("Yeni Kayıt Ekle butonuna tıklandı.")
                    
                    # Kategori dropdown'ından 632 değerini seç
                    print("Kategori dropdown'ından 632 değeri seçiliyor...")
                    drv.execute_script("""
                        var $dropdown = $("input[data-role='dropdownlist']");
                        if ($dropdown.length > 0) {
                            var dropdownlist = $dropdown.data("kendoDropDownList");
                            if (dropdownlist) {
                                dropdownlist.value(632);
                                dropdownlist.trigger('change');
                            }
                        }
                    """)
                    print("Kategori 632 seçildi.")
                    
                    # "Güncelle" butonuna tıkla
                    print("Güncelle butonu aranıyor...")
                    update_button = WebDriverWait(drv, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "a.k-button.k-button-icontext.k-grid-update"))
                    )
                    update_button.click()
                    print("Güncelle butonuna tıklandı.")
                    
                    # Kısa bekleme

                    
                except Exception as e:
                    print(f"Kategori / Marka işlemlerinde hata: {e}")
                    # Hata olsa bile devam et
                
                # "Ürün Varyasyonları" sekmesine tıkla
                print("Ürün Varyasyonları sekmesi aranıyor...")
                try:
                    variations_tab = WebDriverWait(drv, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//li[@data-tab-name='tab-product-attributes']"))
                    )
                    variations_tab.click()
                    print("Ürün Varyasyonları sekmesi tıklandı.")
                except Exception as e:
                    print(f"Sekme tıklama hatası: {e}")
                    continue
                
                                                 # Kombinasyon tablosunu bekle - daha uzun süre bekle
                print("Kombinasyon tablosu yüklenmesi bekleniyor...")
                WebDriverWait(drv, 20).until(
                    EC.presence_of_element_located((By.XPATH, "//tbody[@role='rowgroup']//tr"))
                )
                

                
                # Tüm satırları bul
                rows = drv.find_elements(By.XPATH, "//tbody[@role='rowgroup']//tr")
                print(f"Toplam {len(rows)} satır bulundu.")
                
                # "Beden: S" veya "Beden: 36" olan satırı bul
                target_row = None
                for row in rows:
                    try:
                        # Kombinasyon hücresini bul (2. sütun - Kombinasyon)
                        combination_cell = row.find_elements(By.TAG_NAME, "td")[1]  # 0-based index
                        combination_text = combination_cell.text.strip()
                        
                        print(f"Kombinasyon kontrol ediliyor: '{combination_text}'")
                        
                        if combination_text in ["Beden: S", "Beden: 36"]:
                            target_row = row
                            print(f"Hedef kombinasyon bulundu: {combination_text}")
                            break
                    except Exception as e:
                        print(f"Satır kontrol edilirken hata: {e}")
                        continue
                
                if not target_row:
                    print("'Beden: S' veya 'Beden: 36' olan kombinasyon bulunamadı.")
                    # Debug için tüm satırları yazdır
                    print("Mevcut kombinasyonlar:")
                    for i, row in enumerate(rows):
                        try:
                            combination_cell = row.find_elements(By.TAG_NAME, "td")[1]
                            print(f"Satır {i+1}: {combination_cell.text.strip()}")
                        except:
                            pass
                    continue
                
                # Düzenle butonunu bul ve tıkla
                print("Düzenle butonu aranıyor...")
                try:
                    # Önce direkt button olarak dene
                    edit_button = target_row.find_element(By.XPATH, ".//button[contains(@onclick, 'EditAttributeCombinationPopup')]")
                    print("Düzenle butonu bulundu (onclick ile)")
                except:
                    try:
                        # Alternatif olarak sadece button olarak dene
                        edit_button = target_row.find_element(By.XPATH, ".//button")
                        print("Düzenle butonu bulundu (genel button olarak)")
                    except:
                        try:
                            # Son olarak onclick içeren herhangi bir element olarak dene
                            edit_button = target_row.find_element(By.XPATH, ".//*[contains(@onclick, 'EditAttributeCombinationPopup')]")
                            print("Düzenle butonu bulundu (onclick içeren element olarak)")
                        except Exception as e:
                            print(f"Düzenle butonu bulunamadı: {e}")
                            # Debug için satır içeriğini yazdır
                            print("Hedef satır içeriği:")
                            print(target_row.get_attribute('outerHTML'))
                            continue
                
                # Butona tıkla
                edit_button.click()
                print("Düzenle butonuna tıklandı.")
                
                # Düzenle butonunun onclick'inden kombinasyon ID'sini al
                onclick_value = edit_button.get_attribute("onclick")
                match = re.search(r'/EditAttributeCombinationPopup/(\d+)', onclick_value)
                
                if match:
                    combination_id = match.group(1)
                    print(f"Kombinasyon ID bulundu: {combination_id}")
                    
                    # Direkt popup URL'sine git
                    popup_url = f"{BASE_URL}/admin/product/editattributecombinationpopup/{combination_id}/?btnId=btnRefresh&formId=product-form"
                    print(f"Popup URL'sine gidiliyor: {popup_url}")
                    drv.get(popup_url)
                    
                    # Fiyat alanını bul (Kendo UI numeric textbox için)
                    print("Fiyat alanı aranıyor...")
                    price_input = WebDriverWait(drv, 10).until(
                        EC.presence_of_element_located((By.ID, "OverriddenPrice"))
                    )
                    print("Fiyat alanı bulundu!")
                    
                    # Yeni fiyatı ayarla
                    print(f"Yeni fiyat ayarlanıyor: {variant_price}")
                    try:
                        # Kendo UI numeric textbox için JavaScript ile değeri ayarla
                        drv.execute_script(f"""
                            var numericTextBox = $("#OverriddenPrice").data("kendoNumericTextBox");
                            if (numericTextBox) {{
                                numericTextBox.value({variant_price});
                                // Görünür input'u da güncelle
                                $("#OverriddenPrice + span input.k-formatted-value").val("{variant_price}");
                            }}
                        """)
                        print("JavaScript ile fiyat ayarlandı.")
                    except Exception as js_error:
                        print(f"JavaScript hatası: {js_error}")
                        # Alternatif: Görünür input alanını bul ve güncelle
                        try:
                            visible_input = drv.find_element(By.CSS_SELECTOR, "#OverriddenPrice + span input.k-formatted-value")
                            visible_input.clear()
                            visible_input.send_keys(variant_price)
                            # Hidden input'u da güncelle
                            price_input.clear()
                            price_input.send_keys(variant_price)
                            print("Görünür ve hidden input ile fiyat güncellendi.")
                        except Exception as alt_error:
                            print(f"Alternatif yöntem de başarısız: {alt_error}")
                            # Son çare: Sadece hidden input'u güncelle
                            price_input.clear()
                            price_input.send_keys(variant_price)
                            print("Hidden input ile fiyat güncellendi.")
                    
                    print("Fiyat alanı güncellendi.")
                    
                    # Kaydet butonuna tıkla
                    try:
                        save_button = WebDriverWait(drv, 10).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, "button[name='save']"))
                        )
                        save_button.click()
                        print("Kaydet butonuna tıklandı!")
                        print(f"Fiyat başarıyla güncellendi: {variant_price}")
                        successful_count += 1
                    except Exception as e:
                        print(f"Kaydet butonu bulunamadı: {e}")
                        continue
                    
                else:
                    print("Kombinasyon ID bulunamadı!")
                    continue
                    
            except Exception as e:
                print(f"Düzenle butonu işlenirken hata: {e}")
                continue
                

                
            except Exception as e:
                print(f"Satır {index + 1} işlenirken hata: {e}")
                continue
        
        print(f"\n=== EXCEL GÜNCELLEME TAMAMLANDI ===")
        print(f"Toplam satır: {total_count}")
        print(f"Başarılı: {successful_count}")
        print(f"Başarısız: {total_count - successful_count}")
        
        return True
        
    except Exception as e:
        print(f"Excel güncelleme işlemlerinde hata: {e}")
        return False

def process_product(drv, product_id):
    """Tek bir ürünü işler."""
    try:
        print(f"Ürün {product_id} işleniyor...")
        
        # Ürün düzenleme sayfasına git
        edit_url = f"{BASE_URL}/admin/product/edit/{product_id}"
        drv.get(edit_url)
        
        # Sayfa yüklenmesini bekle
        WebDriverWait(drv, 15).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
        
        # "Ürün Varyasyonları" sekmesine tıkla
        print("Ürün Varyasyonları sekmesi aranıyor...")
        
        # Önce li elementi olarak dene
        try:
            variations_tab = WebDriverWait(drv, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//li[@data-tab-name='tab-product-attributes']//span[contains(text(), 'Ürün Varyasyonları')]"))
            )
            print("Ürün Varyasyonları sekmesi bulundu (li elementi)")
        except:
            # Alternatif olarak direkt span olarak dene
            try:
                variations_tab = WebDriverWait(drv, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Ürün Varyasyonları')]"))
                )
                print("Ürün Varyasyonları sekmesi bulundu (span elementi)")
            except:
                # Son olarak data-tab-name ile dene
                variations_tab = WebDriverWait(drv, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//li[@data-tab-name='tab-product-attributes']"))
                )
                print("Ürün Varyasyonları sekmesi bulundu (data-tab-name ile)")
        
        # Sekmeye tıkla
        variations_tab.click()
        
        # Kombinasyon fiyatı olan satırları bul
        print("Kombinasyon fiyatları kontrol ediliyor...")
        
        # Önce tüm satırları bul
        rows = drv.find_elements(By.XPATH, "//tbody[@role='rowgroup']//tr")
        print(f"Toplam {len(rows)} satır bulundu.")
        
        # Fiyatı dolu olan satırları filtrele
        price_rows = []
        for i, row in enumerate(rows):
            try:
                # Fiyat hücresini bul (6. sütun - Kombinasyon Fiyatı)
                price_cell = row.find_elements(By.TAG_NAME, "td")[5]  # 0-based index
                price_text = price_cell.text.strip()
                
                # Fiyat dolu mu kontrol et
                if price_text and price_text != "" and price_text != "0" and price_text != "0,0000":
                    print(f"Satır {i+1}: Fiyat bulundu: {price_text}")
                    
                    # Bu satırdaki düzenle butonunu bul
                    edit_button = row.find_element(By.XPATH, ".//button[contains(@onclick, 'EditAttributeCombinationPopup')]")
                    price_rows.append((edit_button, price_text))
                else:
                    print(f"Satır {i+1}: Fiyat boş veya 0")
                    
            except Exception as e:
                print(f"Satır {i+1} kontrol edilirken hata: {e}")
                continue
        
        if not price_rows:
            print("Fiyatı dolu olan kombinasyon bulunamadı.")
            return False
        
        print(f"Fiyatı dolu olan {len(price_rows)} kombinasyon bulundu.")
        
        # Her düzenleme butonuna tıkla ve fiyatı sıfırla
        for i, (edit_button, original_price) in enumerate(price_rows):
            try:
                print(f"Kombinasyon {i+1} düzenleniyor... (Mevcut fiyat: {original_price})")
                
                # Düzenle butonuna tıkla
                edit_button.click()
                
                # Düzenle butonunun onclick'inden kombinasyon ID'sini al
                onclick_value = edit_button.get_attribute("onclick")
                print(f"Düzenle butonu onclick: {onclick_value}")
                
                # Kombinasyon ID'sini çıkar
                import re
                match = re.search(r'/EditAttributeCombinationPopup/(\d+)', onclick_value)
                if match:
                    combination_id = match.group(1)
                    print(f"Kombinasyon ID bulundu: {combination_id}")
                    
                    # Direkt popup URL'sine git
                    popup_url = f"{BASE_URL}/admin/product/editattributecombinationpopup/{combination_id}/?btnId=btnRefresh&formId=product-form"
                    print(f"Popup URL'sine gidiliyor: {popup_url}")
                    
                    drv.get(popup_url)
                    
                    # Fiyat alanını bul (Kendo UI numeric textbox için)
                    print("Fiyat alanı aranıyor...")
                    try:
                        # Önce gizli input alanını bul
                        price_input = WebDriverWait(drv, 10).until(
                            EC.presence_of_element_located((By.ID, "OverriddenPrice"))
                        )
                        print("Fiyat alanı bulundu!")
                        
                        # Mevcut değeri kontrol et
                        current_value = price_input.get_attribute("value")
                        print(f"Mevcut fiyat değeri: '{current_value}'")
                        
                        if current_value and current_value != "" and current_value != "0" and current_value != "0,0000" and current_value != "null":
                            print(f"Fiyat siliniyor: {current_value}")
                            
                            # Kendo UI numeric textbox için JavaScript ile değeri tamamen temizle
                            try:
                                drv.execute_script("""
                                    var numericTextBox = $("#OverriddenPrice").data("kendoNumericTextBox");
                                    if (numericTextBox) {
                                        // Değeri tamamen temizle (0 yapma)
                                        numericTextBox.value(null);
                                        // Görünür input'u da temizle
                                        $("#OverriddenPrice + span input.k-formatted-value").val("");
                                    }
                                """)
                                print("JavaScript ile fiyat tamamen temizlendi.")
                            except Exception as js_error:
                                print(f"JavaScript hatası: {js_error}")
                                # Alternatif: Görünür input alanını bul ve tamamen temizle
                                try:
                                    visible_input = drv.find_element(By.CSS_SELECTOR, "#OverriddenPrice + span input.k-formatted-value")
                                    visible_input.clear()
                                    # Hidden input'u da temizle
                                    price_input.clear()
                                    print("Görünür ve hidden input ile fiyat temizlendi.")
                                except Exception as alt_error:
                                    print(f"Alternatif yöntem de başarısız: {alt_error}")
                                    # Son çare: Sadece hidden input'u temizle
                                    price_input.clear()
                                    print("Hidden input ile fiyat temizlendi.")
                            
                            print("Fiyat alanı işlendi.")
                            
                            # Kaydet butonuna tıkla
                            try:
                                save_button = WebDriverWait(drv, 10).until(
                                    EC.element_to_be_clickable((By.CSS_SELECTOR, "button[name='save']"))
                                )
                                save_button.click()
                                print("Kaydet butonuna tıklandı!")
                                print("Fiyat başarıyla silindi!")
                            except Exception as e:
                                print(f"Kaydet butonu bulunamadı: {e}")
                                continue
                        else:
                            print("Fiyat zaten boş veya 0, değişiklik yapılmadı.")
                        
                        # Ürün sayfasına geri dön
                        drv.get(edit_url)
                        
                        # Ürün Varyasyonları sekmesine tekrar tıkla
                        try:
                            variations_tab = WebDriverWait(drv, 10).until(
                                EC.element_to_be_clickable((By.XPATH, "//li[@data-tab-name='tab-product-attributes']"))
                            )
                            variations_tab.click()
                        except:
                            pass
                        
                    except Exception as e:
                        print(f"Fiyat alanı bulunamadı: {e}")
                        # Ürün sayfasına geri dön
                        drv.get(edit_url)
                        continue
                        
                else:
                    print("Kombinasyon ID bulunamadı!")
                    continue
                    
            except Exception as e:
                print(f"Kombinasyon {i+1} işlenirken hata: {e}")
                # Ana pencereye geri dön
                if len(drv.window_handles) > 1:
                    drv.switch_to.window(drv.window_handles[0])
        
        print(f"Ürün {product_id} başarıyla işlendi!")
        return True
        
    except Exception as e:
        print(f"Ürün {product_id} işlenirken hata: {e}")
        return False

def main():
    """Ana fonksiyon."""
    print("XML Otomasyon Programı Başlatılıyor...")
    
    # XML'den ürün ID'lerini al
    product_ids = get_xml_data()
    if not product_ids:
        print("Ürün ID'leri alınamadı. Program sonlandırılıyor.")
        return
    
    # WebDriver'ı başlat
    driver = init_driver()
    if not driver:
        print("WebDriver başlatılamadı. Program sonlandırılıyor.")
        return
    
    try:
        # Sisteme giriş yap
        if not login(driver):
            print("Giriş yapılamadı. Program sonlandırılıyor.")
            return
        
        # Her ürünü işle
        successful_count = 0
        total_count = len(product_ids)
        
        for i, product_id in enumerate(product_ids, 1):
            print(f"\n--- Ürün {i}/{total_count} ---")
            
            if process_product(driver, product_id):
                successful_count += 1
            
            # Ürünler arası çok kısa bekleme (sadece sistem yükünü azaltmak için)
            time.sleep(0.5)
        
        print(f"\n=== İŞLEM TAMAMLANDI ===")
        print(f"Toplam ürün: {total_count}")
        print(f"Başarılı: {successful_count}")
        print(f"Başarısız: {total_count - successful_count}")
        
        # Bulk edit son işlemleri
        print("\nBulk edit son işlemleri başlatılıyor...")
        if bulk_edit_final_operations(driver):
            print("Bulk edit işlemleri başarıyla tamamlandı!")
            
            # Excel'den kombinasyon fiyatlarını güncelle
            print("\nExcel'den kombinasyon fiyatları güncelleniyor...")
            if update_combination_prices_from_excel(driver):
                print("Tüm işlemler başarıyla tamamlandı!")
            else:
                print("Excel güncelleme işlemlerinde hata oluştu!")
        else:
            print("Bulk edit işlemlerinde hata oluştu!")
        
    except KeyboardInterrupt:
        print("\nProgram kullanıcı tarafından durduruldu.")
    except Exception as e:
        print(f"Beklenmeyen hata: {e}")
    finally:
        # Tarayıcıyı kapat
        print("Tarayıcı kapatılıyor...")
        driver.quit()
        print("Program sonlandırıldı.")

if __name__ == "__main__":
    main()





