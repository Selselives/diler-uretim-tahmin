#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Aug 11 08:26:39 2026

@author: selinaysancak
"""
import streamlit as st
import pandas as pd
import joblib
import smtplib

from io import BytesIO
from email.message import EmailMessage
from datetime import date, time, datetime, timedelta


# ============================================================
# SAYFA AYARLARI
# ============================================================

st.set_page_config(
    page_title="Diler | Üretim Süresi Tahmin Sistemi",
    page_icon="Diler_Logo_duzeltilmis.png",
    layout="wide"
)


# ============================================================
# TASARIM
# ============================================================

st.markdown(
    """
    <style>

    .stApp {
        background-color: #ffffff;
    }

    .block-container {
        max-width: 1200px;
        padding-top: 40px;
    }

    h1 {
        color: #1239A6;
        font-size: 46px !important;
        font-weight: 700 !important;
    }

    h2, h3 {
        color: #172033;
    }

    .subtitle {
        color: #6B7280;
        font-size: 18px;
        margin-bottom: 35px;
    }

    .card {
        background-color: #F4F6FB;
        border: 1px solid #E1E5EF;
        border-radius: 15px;
        padding: 22px;
        margin-bottom: 20px;
    }

    .card-title {
        color: #6B7280;
        font-size: 15px;
    }

    .card-value {
        color: #1239A6;
        font-size: 30px;
        font-weight: 700;
    }

    .total-box {
        background-color: #1239A6;
        color: white;
        border-radius: 15px;
        padding: 25px;
        margin-top: 20px;
        margin-bottom: 25px;
    }

    .total-title {
        font-size: 15px;
    }

    .total-value {
        font-size: 30px;
        font-weight: 700;
    }

    .stDownloadButton button {
        background-color: #1239A6 !important;
        color: white !important;
        border-radius: 25px !important;
        border: none !important;
        padding: 10px 25px !important;
        font-weight: 600 !important;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# LOGO VE BAŞLIK
# ============================================================

st.image(
    "Diler_Logo_duzeltilmis.png",
    width=300
)

st.title(
    "Üretim Süresi Tahmin Sistemi"
)

st.markdown(
    """
    <div class="subtitle">
    SAP üretim verileri kullanılarak ürün bazında tahmini üretim sürelerinin hesaplanması
    </div>
    """,
    unsafe_allow_html=True
)

st.divider()


# ============================================================
# MODELİ YÜKLE
# ============================================================

model = joblib.load(
    "model.pkl"
)

model_columns = joblib.load(
    "model_columns.pkl"
)


# ============================================================
# EĞİTİM VERİSİNİ YÜKLE
# ============================================================

try:

    egitim_verisi = pd.read_excel(
        "Diler Proje Verileri.xlsx"
    )

except FileNotFoundError:

    st.error(
        "Diler Proje Verileri.xlsx dosyası bulunamadı. "
        "Bu dosyanın GitHub repository içerisinde app.py ile "
        "aynı klasörde bulunması gerekiyor."
    )

    st.stop()


# ============================================================
# EĞİTİM VERİSİ SÜTUN KONTROLÜ
# ============================================================

egitim_gerekli_sutunlar = [
    "Y_CAP_FLM_MM",
    "Y_KALITE_FLM",
    "Y_KALITE_KTK"
]

egitim_eksik = [
    sutun
    for sutun in egitim_gerekli_sutunlar
    if sutun not in egitim_verisi.columns
]

if egitim_eksik:

    st.error(
        "Diler Proje Verileri.xlsx dosyasında şu sütunlar eksik: "
        + ", ".join(egitim_eksik)
    )

    st.stop()


# ============================================================
# KARŞILAŞTIRMA İÇİN TEMİZLEME FONKSİYONLARI
# ============================================================

def temiz_deger(deger):

    if pd.isna(deger):
        return None

    return str(deger).strip().upper()


def temiz_cap(deger):

    if pd.isna(deger):
        return None

    try:
        return round(float(deger), 6)

    except:
        return None


# ============================================================
# EĞİTİM VERİSİNDE BULUNAN GEÇERLİ KOMBİNASYONLAR
# ============================================================

gecerli_kombinasyonlar = set()

for _, satir in egitim_verisi.iterrows():

    kombinasyon = (
        temiz_cap(
            satir["Y_CAP_FLM_MM"]
        ),

        temiz_deger(
            satir["Y_KALITE_FLM"]
        ),

        temiz_deger(
            satir["Y_KALITE_KTK"]
        )
    )

    gecerli_kombinasyonlar.add(
        kombinasyon
    )


# ============================================================
# SAP ÜRÜN LİSTESİ
# ============================================================

st.subheader(
    "SAP Ürün Listesi"
)

st.write(
    "SAP sisteminden alınan Excel dosyasını yükleyerek tahmin işlemini başlatın."
)


dosya = st.file_uploader(
    "Excel dosyasını seçin",
    type=["xlsx"],
    label_visibility="collapsed"
)


# ============================================================
# EXCEL YÜKLENDİYSE
# ============================================================

if dosya is not None:

    df = pd.read_excel(
        dosya
    )


    # ========================================================
    # SAP EXCELİNDE OLMASI GEREKEN SÜTUNLAR
    # ========================================================

    gerekli_sutunlar = [
        "Filmasin Cap mm -FILMASIN",
        "Mamul Kalitesi -FILMASIN",
        "Uretilecek Paket Sayisi -FILMASIN",
        "Kutuk Kalitesi -KUTUK"
    ]


    eksik = [
        sutun
        for sutun in gerekli_sutunlar
        if sutun not in df.columns
    ]


    if eksik:

        st.error(
            "Excel dosyasında şu sütunlar eksik: "
            + ", ".join(eksik)
        )

        st.stop()


    st.success(
        "Excel başarıyla yüklendi."
    )


    # ========================================================
    # BİLGİ KARTLARI
    # ========================================================

    col1, col2 = st.columns(2)


    with col1:

        st.markdown(
            f"""
            <div class="card">

                <div class="card-title">
                    Ürün / Parti Sayısı
                </div>

                <div class="card-value">
                    {len(df):,}
                </div>

            </div>
            """,
            unsafe_allow_html=True
        )


    # ========================================================
    # SADECE GEREKLİ SÜTUNLARI AL
    # ========================================================

    yeni = df[
        [
            "Filmasin Cap mm -FILMASIN",
            "Mamul Kalitesi -FILMASIN",
            "Uretilecek Paket Sayisi -FILMASIN",
            "Kutuk Kalitesi -KUTUK"
        ]
    ].copy()


    # ========================================================
    # MİKTARI SAYISAL YAP
    # ========================================================

    yeni[
        "Uretilecek Paket Sayisi -FILMASIN"
    ] = pd.to_numeric(
        yeni[
            "Uretilecek Paket Sayisi -FILMASIN"
        ],
        errors="coerce"
    )


    # ========================================================
    # MODEL İÇİN SÜTUNLARI ESKİ İSİMLERE ÇEVİR
    # ========================================================

    model_input = yeni.rename(
        columns={
            "Filmasin Cap mm -FILMASIN":
                "Y_CAP_FLM_MM",

            "Mamul Kalitesi -FILMASIN":
                "Y_KALITE_FLM",

            "Uretilecek Paket Sayisi -FILMASIN":
                "Miktar",

            "Kutuk Kalitesi -KUTUK":
                "Y_KALITE_KTK"
        }
    )


    # ========================================================
    # MODEL VERİSİ OLUŞTUR
    # ========================================================

    model_data = pd.get_dummies(
        model_input[
            [
                "Y_CAP_FLM_MM",
                "Y_KALITE_FLM",
                "Y_KALITE_KTK"
            ]
        ],
        columns=[
            "Y_KALITE_FLM",
            "Y_KALITE_KTK"
        ]
    )


    model_data = model_data.reindex(
        columns=model_columns,
        fill_value=0
    )


    # ========================================================
    # EĞİTİM VERİSİNDE VAR MI KONTROLÜ
    # ========================================================

    gecerli_durumlar = []


    for _, satir in yeni.iterrows():

        kombinasyon = (
            temiz_cap(
                satir[
                    "Filmasin Cap mm -FILMASIN"
                ]
            ),

            temiz_deger(
                satir[
                    "Mamul Kalitesi -FILMASIN"
                ]
            ),

            temiz_deger(
                satir[
                    "Kutuk Kalitesi -KUTUK"
                ]
            )
        )


        gecerli_durumlar.append(
            kombinasyon in gecerli_kombinasyonlar
        )


    yeni[
        "TAHMIN_YAPILABILIR"
    ] = gecerli_durumlar


    # ========================================================
    # TAHMİN SÜTUNLARINI OLUŞTUR
    # ========================================================

    yeni[
        "TAHMINI_1_KUTUK_SURESI"
    ] = pd.Series(
        [None] * len(yeni),
        index=yeni.index,
        dtype="object"
    )


    yeni[
        "TAHMINI_TOPLAM_SURE"
    ] = 0.0


    yeni[
        "TAHMIN_DURUMU"
    ] = "Tahmin yapılamadı"


    # ========================================================
    # SADECE GEÇERLİ SATIRLARI TAHMİN ET
    # ========================================================

    gecerli_pozisyonlar = [
        i
        for i, durum in enumerate(
            yeni[
                "TAHMIN_YAPILABILIR"
            ]
        )
        if durum
    ]


    if len(gecerli_pozisyonlar) > 0:

        model_tahmin_data = model_data.iloc[
            gecerli_pozisyonlar
        ]


        tahmin = model.predict(
            model_tahmin_data
        )


        tahmin = tahmin.flatten()


        # ====================================================
        # TAHMİNLERİ SATIRLARA YAZ
        # ====================================================

        for sira, pozisyon in enumerate(
            gecerli_pozisyonlar
        ):

            index = yeni.index[
                pozisyon
            ]


            bir_kutuk_suresi = float(
                tahmin[sira]
            )


            miktar = float(
                yeni.iloc[
                    pozisyon
                ][
                    "Uretilecek Paket Sayisi -FILMASIN"
                ]
            )


            toplam_satir_suresi = (
                bir_kutuk_suresi
                * miktar
            )


            yeni.at[
                index,
                "TAHMINI_1_KUTUK_SURESI"
            ] = round(
                bir_kutuk_suresi,
                2
            )


            yeni.at[
                index,
                "TAHMINI_TOPLAM_SURE"
            ] = round(
                toplam_satir_suresi,
                2
            )


            yeni.at[
                index,
                "TAHMIN_DURUMU"
            ] = "Tahmin oluşturuldu"


    # ========================================================
    # TOPLAM SÜRE
    # ========================================================

    toplam_sure = float(
        yeni[
            "TAHMINI_TOPLAM_SURE"
        ].sum()
    )


    saat = int(
        toplam_sure // 3600
    )


    dakika = int(
        (toplam_sure % 3600) // 60
    )


    saniye = int(
        toplam_sure % 60
    )


    # ========================================================
    # TAHMİN DURUM SAYILARI
    # ========================================================

    tahmin_yapilabilen = int(
        yeni[
            "TAHMIN_YAPILABILIR"
        ].sum()
    )


    tahmin_yapilamayan = (
        len(yeni)
        - tahmin_yapilabilen
    )


    # ========================================================
    # KULLANICIYA BİLGİ
    # ========================================================

    if tahmin_yapilamayan > 0:

        st.warning(
            f"⚠️ {tahmin_yapilamayan} ürün için "
            f"tahmin oluşturulamadı. "
            f"Bu ürünlerin çap + mamul kalitesi + "
            f"kütük kalitesi kombinasyonu "
            f"Diler Proje Verileri eğitim verisinde "
            f"bulunmamaktadır. "
            f"Toplam süre yalnızca tahmin yapılabilen "
            f"{tahmin_yapilabilen} ürün üzerinden hesaplanmıştır."
        )

    else:

        st.success(
            f"✅ {tahmin_yapilabilen} ürünün tamamı "
            f"için tahmin oluşturuldu."
        )


    # ========================================================
    # TOPLAM TAHMİNİ SÜRE KARTI
    # ========================================================

    with col2:

        st.markdown(
            f"""
            <div class="card">

                <div class="card-title">
                    Toplam Tahmini Süre
                </div>

                <div class="card-value">
                    {saat} sa {dakika} dk
                </div>

            </div>
            """,
            unsafe_allow_html=True
        )


    # ========================================================
    # ÜRETİM BAŞLANGIÇ BİLGİLERİ
    # ========================================================

    st.divider()

    st.subheader(
        "Üretim Başlangıç Bilgileri"
    )


    col_tarih, col_saat = st.columns(2)


    with col_tarih:

        baslangic_tarihi = st.date_input(
            "Başlangıç Tarihi",
            value=date.today()
        )


    with col_saat:

        baslangic_saati = st.time_input(
            "Başlangıç Saati",
            value=time(8, 0)
        )


    # ========================================================
    # BİTİŞ ZAMANI
    # ========================================================

    baslangic_zamani = datetime.combine(
        baslangic_tarihi,
        baslangic_saati
    )


    bitis_zamani = (
        baslangic_zamani
        + timedelta(
            seconds=toplam_sure
        )
    )


    # ========================================================
    # BİTİŞ ZAMANI GÖSTER
    # ========================================================

    st.markdown(
        f"""
        <div class="total-box">

            <div class="total-title">
                TAHMİNİ ÜRETİM BİTİŞ ZAMANI
            </div>

            <div class="total-value">
                {bitis_zamani.strftime("%d.%m.%Y %H:%M")}
            </div>

        </div>
        """,
        unsafe_allow_html=True
    )


    # ========================================================
    # TAHMİN SONUÇLARI
    # ========================================================

    st.subheader(
        "Tahmin Sonuçları"
    )


    # Teknik kontrol sütunlarını kullanıcıdan gizle
    sonuc_gosterim = yeni.drop(
        columns=[
            "TAHMIN_YAPILABILIR"
        ]
    )


    st.dataframe(
        sonuc_gosterim,
        use_container_width=True,
        hide_index=True
    )


    # ========================================================
    # TOPLAM SÜRE
    # ========================================================

    st.markdown(
        f"""
        <div class="total-box">

            <div class="total-title">
                TOPLAM TAHMİNİ ÜRETİM SÜRESİ
            </div>

            <div class="total-value">
                {saat} saat {dakika} dakika {saniye} saniye
            </div>

        </div>
        """,
        unsafe_allow_html=True
    )


    # ========================================================
    # EXCEL OLUŞTUR
    # ========================================================

    sonuc_excel = BytesIO()


    with pd.ExcelWriter(
        sonuc_excel,
        engine="openpyxl"
    ) as writer:

        sonuc_gosterim.to_excel(
            writer,
            index=False,
            sheet_name="Tahmin Sonuclari"
        )


    sonuc_excel.seek(0)


    # ========================================================
    # EXCEL İNDİR
    # ========================================================

    st.download_button(
        label="Tahmin Sonuçlarını Excel Olarak İndir",

        data=sonuc_excel.getvalue(),

        file_name="Diler_Tahmin_Sonuclari.xlsx",

        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )


    st.write("")


    # ========================================================
    # E-POSTA GÖNDER
    # ========================================================

    if st.button(
        "📧 E-posta ile Gönder"
    ):

        try:

            email_address = st.secrets[
                "EMAIL_ADDRESS"
            ]

            email_password = st.secrets[
                "EMAIL_PASSWORD"
            ]

            to_email = st.secrets[
                "TO_EMAIL"
            ]


            mesaj = EmailMessage()


            mesaj["Subject"] = (
                "Diler Üretim Süresi Tahmin Sonuçları"
            )

            mesaj["From"] = email_address

            mesaj["To"] = to_email


            mesaj.set_content(
                f"""
Merhaba,

Diler Üretim Süresi Tahmin Sistemi tarafından oluşturulan
tahmin sonuçları ekte paylaşılmıştır.

Toplam tahmini üretim süresi:
{saat} saat {dakika} dakika {saniye} saniye

Tahmin yapılabilen ürün sayısı:
{tahmin_yapilabilen}

Tahmin yapılamayan ürün sayısı:
{tahmin_yapilamayan}

Tahmini üretim bitiş zamanı:
{bitis_zamani.strftime("%d.%m.%Y %H:%M")}

İyi çalışmalar.
"""
            )


            mesaj.add_attachment(
                sonuc_excel.getvalue(),

                maintype="application",

                subtype=(
                    "vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),

                filename=(
                    "Diler_Tahmin_Sonuclari.xlsx"
                )
            )


            with smtplib.SMTP_SSL(
                "smtp.gmail.com",
                465
            ) as server:

                server.login(
                    email_address,
                    email_password
                )

                server.send_message(
                    mesaj
                )


            st.success(
                "E-posta başarıyla gönderildi! ✅"
            )


        except Exception as e:

            st.error(
                "E-posta gönderilemedi. "
                "Secrets ayarlarını kontrol edin."
            )

            st.write(e)