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


# ------------------------------------------------
# SAYFA AYARLARI
# ------------------------------------------------

st.set_page_config(
    page_title="Diler | Üretim Süresi Tahmin Sistemi",
    page_icon="Diler_Logo_duzeltilmis.png",
    layout="wide"
)


# ------------------------------------------------
# TASARIM
# ------------------------------------------------

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


# ------------------------------------------------
# LOGO VE BAŞLIK
# ------------------------------------------------

st.image(
    "Diler_Logo_duzeltilmis.png",
    width=300
)

st.title(
    "Üretim Süresi Tahmin Sistemi"
)

st.markdown(
    '<div class="subtitle">'
    'SAP üretim verileri kullanılarak ürün bazında tahmini üretim sürelerinin hesaplanması'
    '</div>',
    unsafe_allow_html=True
)

st.divider()


# ------------------------------------------------
# MODEL
# ------------------------------------------------

model = joblib.load("model.pkl")
model_columns = joblib.load("model_columns.pkl")
# ------------------------------------------------
# EĞİTİM VERİSİNİ YÜKLE
# ------------------------------------------------

egitim_verisi = pd.read_excel(
    "Diler Proje Verileri.xlsx"
)

gecerli_kombinasyonlar = set(
    zip(
        egitim_verisi["Y_CAP_FLM_MM"],
        egitim_verisi["Y_KALITE_FLM"],
        egitim_verisi["Y_KALITE_KTK"]
    )
)


# ------------------------------------------------
# SAP EXCEL YÜKLEME
# ------------------------------------------------

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


# ------------------------------------------------
# EXCEL YÜKLENDİYSE
# ------------------------------------------------

if dosya is not None:

    df = pd.read_excel(dosya)


    # ------------------------------------------------
    # YENİ SAP SÜTUNLARI
    # ------------------------------------------------

    gerekli_sutunlar = [
        "Filmasin Cap mm -FILMASIN",
        "Mamul Kalitesi -FILMASIN",
        "Uretilecek Paket Sayisi -FILMASIN",
        "Kutuk Kalitesi -KUTUK"
    ]


    # ------------------------------------------------
    # EKSİK SÜTUN KONTROLÜ
    # ------------------------------------------------

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


    else:

        st.success(
            "Excel başarıyla yüklendi."
        )


        # ------------------------------------------------
        # BİLGİ KARTLARI
        # ------------------------------------------------

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


        # ------------------------------------------------
        # KULLANICIYA GÖSTERİLECEK VERİ
        # ------------------------------------------------

        yeni = df[
            [
                "Filmasin Cap mm -FILMASIN",
                "Mamul Kalitesi -FILMASIN",
                "Uretilecek Paket Sayisi -FILMASIN",
                "Kutuk Kalitesi -KUTUK"
            ]
        ].copy()


        # ------------------------------------------------
        # MİKTARI SAYISAL YAP
        # ------------------------------------------------

        yeni[
            "Uretilecek Paket Sayisi -FILMASIN"
        ] = pd.to_numeric(
            yeni[
                "Uretilecek Paket Sayisi -FILMASIN"
            ],
            errors="coerce"
        )


        # ------------------------------------------------
        # MODEL İÇİN VERİLERİ ESKİ İSİMLERE ÇEVİR
        # ------------------------------------------------

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


        # ------------------------------------------------
        # MODEL VERİSİ
        # ------------------------------------------------

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
        # ------------------------------------------------
        # EĞİTİM VERİSİNDE BULUNAN KOMBİNASYONLARI KONTROL ET
        # ------------------------------------------------

gecerli_satirlar = []

for _, satir in yeni.iterrows():

    kombinasyon = (
        satir["Filmasin Cap mm -FILMASIN"],
        satir["Mamul Kalitesi -FILMASIN"],
        satir["Kutuk Kalitesi -KUTUK"]
    )

    gecerli_satirlar.append(
        kombinasyon in gecerli_kombinasyonlar
    )

yeni["TAHMIN_YAPILABILIR"] = gecerli_satirlar


# ------------------------------------------------
# TAHMİN SÜTUNLARINI OLUŞTUR
# ------------------------------------------------

yeni["TAHMINI_1_KUTUK_SURESI"] = "Tahmin yapılamadı"

yeni["TAHMINI_TOPLAM_SURE"] = 0.0


# ------------------------------------------------
# SADECE GEÇERLİ KOMBİNASYONLARI TAHMİN ET
# ------------------------------------------------

gecerli_index = yeni.index[
    yeni["TAHMIN_YAPILABILIR"]
]


if len(gecerli_index) > 0:

    model_tahmin_data = model_data.loc[
        gecerli_index
    ]

    tahmin = model.predict(
        model_tahmin_data
    )

    tahmin = tahmin.flatten()


    for i, index in enumerate(gecerli_index):

        yeni.at[
            index,
            "TAHMINI_1_KUTUK_SURESI"
        ] = round(
            float(tahmin[i]),
            2
        )

        yeni.at[
            index,
            "TAHMINI_TOPLAM_SURE"
        ] = round(
            float(tahmin[i])
            * float(
                yeni.at[
                    index,
                    "Uretilecek Paket Sayisi -FILMASIN"
                ]
            ),
            2
        )


# ------------------------------------------------
# TOPLAM SÜRE
# ------------------------------------------------

toplam_sure = (
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


# ------------------------------------------------
# TAHMİN DURUMU
# ------------------------------------------------

tahmin_yapilamayan = (
    (~yeni["TAHMIN_YAPILABILIR"]).sum()
)

tahmin_yapilabilen = (
    yeni["TAHMIN_YAPILABILIR"].sum()
)


if tahmin_yapilamayan > 0:

    st.warning(
        f"⚠️ {tahmin_yapilamayan} ürün için tahmin "
        f"oluşturulamadı. Bu ürünlerin çap / mamul kalitesi / "
        f"kütük kalitesi kombinasyonları eğitim verilerinde "
        f"bulunmamaktadır. Toplam süre yalnızca tahmin "
        f"yapılabilen {tahmin_yapilabilen} ürün üzerinden "
        f"hesaplanmıştır."
    )


    # ------------------------------------------------
    # EĞİTİM VERİSİNDE BULUNAN KOMBİNASYONLARI KONTROL ET
    # ------------------------------------------------

    gecerli_satirlar = []

    for _, satir in yeni.iterrows():

        kombinasyon = (
            satir["Filmasin Cap mm -FILMASIN"],
            satir["Mamul Kalitesi -FILMASIN"],
            satir["Kutuk Kalitesi -KUTUK"]
        )

        gecerli_satirlar.append(
            kombinasyon in gecerli_kombinasyonlar
        )

    yeni["TAHMIN_YAPILABILIR"] = gecerli_satirlar


    # ------------------------------------------------
    # TAHMİN SÜTUNLARINI OLUŞTUR
    # ------------------------------------------------

    yeni["TAHMINI_1_KUTUK_SURESI"] = "Tahmin yapılamadı"

    yeni["TAHMINI_TOPLAM_SURE"] = 0.0


# ------------------------------------------------
# SADECE GEÇERLİ KOMBİNASYONLARI TAHMİN ET
# ------------------------------------------------

gecerli_index = yeni.index[
    yeni["TAHMIN_YAPILABILIR"]
]


if len(gecerli_index) > 0:

    model_tahmin_data = model_data.loc[
        gecerli_index
    ]

    tahmin = model.predict(
        model_tahmin_data
    )

    # Tahmin sonucunu tek boyutlu hale getir
    tahmin = tahmin.flatten()

    # Tahminleri güvenli şekilde yerleştir
    for i, index in enumerate(gecerli_index):

        yeni.at[
            index,
            "TAHMINI_1_KUTUK_SURESI"
        ] = round(
            float(tahmin[i]),
            2
        )

        yeni.at[
            index,
            "TAHMINI_TOPLAM_SURE"
        ] = round(
            float(tahmin[i])
            * float(
                yeni.at[
                    index,
                    "Uretilecek Paket Sayisi -FILMASIN"
                ]
            ),
            2
        )


    # ------------------------------------------------
    # TOPLAM SÜRE
    # ------------------------------------------------

    toplam_sure = (
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


    # ------------------------------------------------
    # TAHMİN DURUMU
    # ------------------------------------------------

    tahmin_yapilamayan = (
        (~yeni["TAHMIN_YAPILABILIR"]).sum()
    )

    tahmin_yapilabilen = (
        yeni["TAHMIN_YAPILABILIR"].sum()
    )


    if tahmin_yapilamayan > 0:

        st.warning(
            f"⚠️ {tahmin_yapilamayan} ürün için tahmin "
            f"oluşturulamadı. Bu ürünlerin çap / mamul kalitesi / "
            f"kütük kalitesi kombinasyonları eğitim verilerinde "
            f"bulunmamaktadır. Toplam süre yalnızca tahmin "
            f"yapılabilen {tahmin_yapilabilen} ürün üzerinden "
            f"hesaplanmıştır."
        )

        # ------------------------------------------------
        # TOPLAM TAHMİNİ SÜRE KARTI
        # ------------------------------------------------

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


        st.divider()


        # ------------------------------------------------
        # ÜRETİM BAŞLANGIÇ BİLGİLERİ
        # ------------------------------------------------

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


        # ------------------------------------------------
        # BİTİŞ ZAMANI HESAPLA
        # ------------------------------------------------

        baslangic_zamani = datetime.combine(
            baslangic_tarihi,
            baslangic_saati
        )


        bitis_zamani = (
            baslangic_zamani
            + timedelta(
                seconds=float(toplam_sure)
            )
        )


        # ------------------------------------------------
        # TAHMİNİ BİTİŞ ZAMANI
        # ------------------------------------------------

        st.subheader("Tahmini Üretim Bitiş Zamanı")

        st.success(
    f"🏭 Üretimin tahmini bitiş zamanı: "
    f"{bitis_zamani.strftime('%d.%m.%Y %H:%M')}"
)

        # ------------------------------------------------
        # TAHMİN SONUÇLARI
        # ------------------------------------------------

        st.subheader(
            "Tahmin Sonuçları"
        )


        sonuc_gosterim = yeni.drop(
    columns=["TAHMIN_YAPILABILIR"]
)

        st.dataframe(
    sonuc_gosterim,
    use_container_width=True,
    hide_index=True
)


        # ------------------------------------------------
        # TOPLAM SÜRE
        # ------------------------------------------------

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


        # ------------------------------------------------
        # EXCEL OLUŞTUR
        # ------------------------------------------------

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


        # ------------------------------------------------
        # EXCEL İNDİR
        # ------------------------------------------------

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


        # ------------------------------------------------
        # E-POSTA GÖNDER
        # ------------------------------------------------

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