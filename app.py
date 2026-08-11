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
from datetime import datetime, date, time, timedelta
from email.message import EmailMessage


# ============================================================
# SAYFA AYARLARI
# ============================================================

st.set_page_config(
    page_title="Diler | Üretim Süresi Tahmin Sistemi",
    page_icon="🔷",
    layout="wide"
)


# ============================================================
# CSS
# ============================================================

st.markdown(
    """
<style>

.stApp {
    background-color: #ffffff;
}

.block-container {
    max-width: 1200px;
    padding-top: 35px;
    padding-bottom: 50px;
}

h1 {
    color: #1239A6 !important;
    font-size: 46px !important;
    font-weight: 700 !important;
}

h2, h3 {
    color: #172033 !important;
}

.stDownloadButton button {
    background-color: #1239A6 !important;
    color: white !important;
    border-radius: 25px !important;
    border: none !important;
    padding: 10px 25px !important;
    font-weight: 600 !important;
}

.stButton button {
    border-radius: 25px !important;
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
    "SAP üretim verileri kullanılarak ürün bazında tahmini üretim sürelerinin hesaplanması"
)

st.divider()


# ============================================================
# MODEL DOSYALARINI YÜKLE
# ============================================================

try:

    model = joblib.load("model.pkl")
    model_columns = joblib.load("model_columns.pkl")

except Exception as e:

    st.error(
        "Model dosyaları yüklenemedi. "
        "model.pkl ve model_columns.pkl dosyalarının GitHub deposunda bulunduğundan emin olun."
    )

    st.stop()


# ============================================================
# EĞİTİM VERİSİNİ YÜKLE
# ============================================================

EGITIM_DOSYASI = "Diler Proje Verileri.xlsx"

try:

    egitim_verisi = pd.read_excel(
        EGITIM_DOSYASI
    )

except FileNotFoundError:

    st.error(
        f"'{EGITIM_DOSYASI}' dosyası bulunamadı. "
        "Bu dosyanın app.py ile aynı GitHub klasöründe olduğundan emin olun."
    )

    st.stop()

except Exception as e:

    st.error(
        f"Eğitim verisi okunamadı: {e}"
    )

    st.stop()


# ============================================================
# EĞİTİM VERİSİ SÜTUNLARINI KONTROL ET
# ============================================================

egitim_gerekli = [
    "Y_CAP_FLM_MM",
    "Y_KALITE_FLM",
    "Y_KALITE_KTK"
]

egitim_eksik = [
    sutun
    for sutun in egitim_gerekli
    if sutun not in egitim_verisi.columns
]

if egitim_eksik:

    st.error(
        "Diler Proje Verileri.xlsx içerisinde gerekli sütunlar bulunamadı: "
        + ", ".join(egitim_eksik)
    )

    st.stop()


# ============================================================
# EĞİTİM VERİSİNDEN GEÇERLİ KOMBİNASYONLARI OLUŞTUR
# ============================================================

def temiz_deger(deger):

    if pd.isna(deger):
        return ""

    return str(deger).strip()


egitim_kombinasyonlari = set()

for _, satir in egitim_verisi.iterrows():

    cap = temiz_deger(
        satir["Y_CAP_FLM_MM"]
    )

    mamul_kalitesi = temiz_deger(
        satir["Y_KALITE_FLM"]
    )

    kutuk_kalitesi = temiz_deger(
        satir["Y_KALITE_KTK"]
    )

    if (
        cap != ""
        and mamul_kalitesi != ""
        and kutuk_kalitesi != ""
    ):

        kombinasyon = (
            cap,
            mamul_kalitesi,
            kutuk_kalitesi
        )

        egitim_kombinasyonlari.add(
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
# DOSYA YÜKLENDİYSE
# ============================================================

if dosya is not None:

    # --------------------------------------------------------
    # EXCEL OKU
    # --------------------------------------------------------

    try:

        df = pd.read_excel(
            dosya
        )

    except Exception as e:

        st.error(
            f"Excel dosyası okunamadı: {e}"
        )

        st.stop()


    # --------------------------------------------------------
    # BEKLENEN SÜTUNLAR
    # --------------------------------------------------------

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
    # KULLANICI VERİSİNİ HAZIRLA
    # ========================================================

    yeni = df[
        gerekli_sutunlar
    ].copy()


    # --------------------------------------------------------
    # MİKTARI SAYISAL YAP
    # --------------------------------------------------------

    yeni[
        "Uretilecek Paket Sayisi -FILMASIN"
    ] = pd.to_numeric(
        yeni[
            "Uretilecek Paket Sayisi -FILMASIN"
        ],
        errors="coerce"
    ).fillna(0)


    # ========================================================
    # MODEL İÇİN ESKİ SÜTUN İSİMLERİNE ÇEVİR
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
    ).copy()


    # ========================================================
    # GEÇERLİ KOMBİNASYON KONTROLÜ
    # ========================================================

    gecerli_maskeler = []

    for _, satir in model_input.iterrows():

        cap = temiz_deger(
            satir["Y_CAP_FLM_MM"]
        )

        mamul_kalitesi = temiz_deger(
            satir["Y_KALITE_FLM"]
        )

        kutuk_kalitesi = temiz_deger(
            satir["Y_KALITE_KTK"]
        )

        kombinasyon = (
            cap,
            mamul_kalitesi,
            kutuk_kalitesi
        )

        gecerli_maskeler.append(
            kombinasyon in egitim_kombinasyonlari
        )


    gecerli_maskesi = pd.Series(
        gecerli_maskeler,
        index=model_input.index
    )


    # ========================================================
    # SONUÇ SÜTUNLARINI BAŞLANGIÇTA OLUŞTUR
    # ========================================================

    yeni[
        "TAHMINI_1_KUTUK_SURESI"
    ] = 0.0

    yeni[
        "TAHMINI_TOPLAM_SURE"
    ] = 0.0

    yeni[
        "TAHMIN_DURUMU"
    ] = "Tahmin yapılamadı"


    # ========================================================
    # SADECE GEÇERLİ KOMBİNASYONLARI TAHMİN ET
    # ========================================================

    if gecerli_maskesi.any():

        gecerli_model_input = model_input.loc[
            gecerli_maskesi
        ].copy()


        # ----------------------------------------------------
        # MODEL VERİSİ
        # ----------------------------------------------------

        model_data = pd.get_dummies(
            gecerli_model_input[
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


        # ----------------------------------------------------
        # MODEL SÜTUNLARIYLA EŞLEŞTİR
        # ----------------------------------------------------

        model_data = model_data.reindex(
            columns=model_columns,
            fill_value=0
        )


        # ----------------------------------------------------
        # TAHMİN
        # ----------------------------------------------------

        try:

            tahmin = model.predict(
                model_data
            )

        except Exception as e:

            st.error(
                f"Model tahmini sırasında hata oluştu: {e}"
            )

            st.stop()


        # ----------------------------------------------------
        # TAHMİN SONUÇLARINI YERİNE YAZ
        # ----------------------------------------------------

        gecerli_index = (
            model_input.index[
                gecerli_maskesi
            ]
        )


        yeni.loc[
            gecerli_index,
            "TAHMINI_1_KUTUK_SURESI"
        ] = pd.Series(
            tahmin,
            index=gecerli_index
        ).round(2)


        # ----------------------------------------------------
        # MİKTARLA ÇARP
        # ----------------------------------------------------

        toplam_sure_serisi = (
            pd.Series(
                tahmin,
                index=gecerli_index
            )
            *
            model_input.loc[
                gecerli_index,
                "Miktar"
            ]
        ).round(2)


        yeni.loc[
            gecerli_index,
            "TAHMINI_TOPLAM_SURE"
        ] = toplam_sure_serisi


        yeni.loc[
            gecerli_index,
            "TAHMIN_DURUMU"
        ] = "Tahmin yapıldı"


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
    # SAYILAR
    # ========================================================

    toplam_urun = len(
        yeni
    )

    tahmin_yapilan = int(
        gecerli_maskesi.sum()
    )

    tahmin_yapilamayan = (
        toplam_urun
        - tahmin_yapilan
    )


    # ========================================================
    # ÖZET KARTLARI
    # ========================================================

    col1, col2, col3 = st.columns(3)


    with col1:

        st.metric(
            label="Toplam Ürün / Parti",
            value=f"{toplam_urun:,}"
        )


    with col2:

        st.metric(
            label="Tahmin Yapılabilen",
            value=f"{tahmin_yapilan:,}"
        )


    with col3:

        st.metric(
            label="Tahmin Yapılamayan",
            value=f"{tahmin_yapilamayan:,}"
        )


    # ========================================================
    # UYARI
    # ========================================================

    if tahmin_yapilamayan > 0:

        st.warning(
            f"{tahmin_yapilamayan} ürün için tahmin oluşturulamadı. "
            "Bu ürünlerin çap + mamul kalitesi + kütük kalitesi "
            "kombinasyonu Diler Proje Verileri eğitim verisinde bulunmamaktadır. "
            "Bu ürünlerin tahmini üretim süresi 0 kabul edilmiştir. "
            f"Toplam süre yalnızca tahmin yapılabilen "
            f"{tahmin_yapilan} ürün üzerinden hesaplanmıştır."
        )


    # ========================================================
    # TOPLAM SÜRE
    # ========================================================

    st.markdown(
        f"""
<div style="
background-color:#1239A6;
color:white;
border-radius:15px;
padding:25px;
margin-top:20px;
margin-bottom:25px;
">
<div style="font-size:15px;">
TOPLAM TAHMİNİ ÜRETİM SÜRESİ
</div>
<div style="font-size:30px;font-weight:700;">
{saat} saat {dakika} dakika {saniye} saniye
</div>
</div>
""",
        unsafe_allow_html=True
    )


    # ========================================================
    # ÜRETİM BAŞLANGIÇ BİLGİLERİ
    # ========================================================

    st.subheader(
        "Üretim Başlangıç Bilgileri"
    )


    tarih_col, saat_col = st.columns(2)


    with tarih_col:

        baslangic_tarihi = st.date_input(
            "Başlangıç Tarihi",
            value=date.today()
        )


    with saat_col:

        baslangic_saati = st.time_input(
            "Başlangıç Saati",
            value=time(8, 0)
        )


    # ========================================================
    # BİTİŞ ZAMANINI HESAPLA
    # ========================================================

    baslangic_datetime = datetime.combine(
        baslangic_tarihi,
        baslangic_saati
    )


    bitis_datetime = (
        baslangic_datetime
        + timedelta(
            seconds=toplam_sure
        )
    )


    bitis_metni = bitis_datetime.strftime(
        "%d.%m.%Y %H:%M"
    )


    # ========================================================
    # BİTİŞ ZAMANI
    # ========================================================

    st.markdown(
        f"""
<div style="
background-color:#1239A6;
color:white;
border-radius:15px;
padding:25px;
margin-top:20px;
margin-bottom:25px;
">
<div style="font-size:15px;">
TAHMİNİ ÜRETİM BİTİŞ ZAMANI
</div>
<div style="font-size:30px;font-weight:700;">
{bitis_metni}
</div>
</div>
""",
        unsafe_allow_html=True
    )


    # ========================================================
    # TAHMİN SONUÇLARI
    # ========================================================

    st.divider()

    st.subheader(
        "Tahmin Sonuçları"
    )


    st.dataframe(
        yeni,
        use_container_width=True,
        hide_index=True
    )


    # ========================================================
    # EXCEL DOSYASI OLUŞTUR
    # ========================================================

    sonuc_excel = BytesIO()


    with pd.ExcelWriter(
        sonuc_excel,
        engine="openpyxl"
    ) as writer:

        yeni.to_excel(
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
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


    # ========================================================
    # E-POSTA
    # ========================================================

    st.write("")


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

            mesaj["From"] = (
                email_address
            )

            mesaj["To"] = (
                to_email
            )


            mesaj.set_content(
                f"""
Merhaba,

Diler Üretim Süresi Tahmin Sistemi tarafından
oluşturulan tahmin sonuçları ekte paylaşılmıştır.

Toplam ürün / parti sayısı:
{toplam_urun}

Tahmin yapılabilen ürün sayısı:
{tahmin_yapilan}

Tahmin yapılamayan ürün sayısı:
{tahmin_yapilamayan}

Toplam tahmini üretim süresi:
{saat} saat {dakika} dakika {saniye} saniye

Üretim başlangıcı:
{baslangic_datetime.strftime("%d.%m.%Y %H:%M")}

Tahmini üretim bitiş zamanı:
{bitis_metni}

Tahmin yapılamayan kombinasyonların üretim süresi
hesaplamaya 0 olarak dahil edilmiştir.

İyi çalışmalar.
"""
            )


            mesaj.add_attachment(
                sonuc_excel.getvalue(),
                maintype="application",
                subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                filename="Diler_Tahmin_Sonuclari.xlsx"
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
                "Streamlit Secrets ayarlarını kontrol edin."
            )

            st.write(
                str(e)
            )