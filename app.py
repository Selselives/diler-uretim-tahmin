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
import os

from io import BytesIO
from email.message import EmailMessage
from datetime import datetime, date, time, timedelta


# ============================================================
# SAYFA AYARLARI
# ============================================================

st.set_page_config(
    page_title="Diler | Üretim Süresi Tahmin Sistemi",
    page_icon="🔷",
    layout="wide"
)


# ============================================================
# BASİT TASARIM
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

st.title("Üretim Süresi Tahmin Sistemi")

st.write(
    "SAP üretim verileri kullanılarak ürün bazında "
    "tahmini üretim sürelerinin hesaplanması"
)

st.divider()


# ============================================================
# MODELİ YÜKLE
# ============================================================

model = joblib.load("model.pkl")
model_columns = joblib.load("model_columns.pkl")


# ============================================================
# EĞİTİM VERİSİNİ YÜKLE
# ============================================================

egitim_dosyasi = "Diler Proje Verileri.xlsx"

if not os.path.exists(egitim_dosyasi):

    st.error(
        f"'{egitim_dosyasi}' dosyası uygulamanın bulunduğu "
        "GitHub klasöründe bulunamadı."
    )

    st.info(
        "Bu dosyayı app.py, model.pkl ve model_columns.pkl "
        "ile aynı klasöre yüklemen gerekiyor."
    )

    st.stop()


egitim_verisi = pd.read_excel(
    egitim_dosyasi
)


# ============================================================
# EĞİTİM VERİSİNDEKİ SÜTUNLARI BELİRLE
# ============================================================

# Yeni sütun isimleri
yeni_sutunlar = [
    "Filmasin Cap mm -FILMASIN",
    "Mamul Kalitesi -FILMASIN",
    "Uretilecek Paket Sayisi -FILMASIN",
    "Kutuk Kalitesi -KUTUK"
]

# Eski/model sütun isimleri
eski_sutunlar = [
    "Y_CAP_FLM_MM",
    "Y_KALITE_FLM",
    "Miktar",
    "Y_KALITE_KTK"
]


# Eğitim dosyasında yeni isimler varsa
if all(
    sutun in egitim_verisi.columns
    for sutun in yeni_sutunlar
):

    egitim_kombinasyonlari = egitim_verisi[
        [
            "Filmasin Cap mm -FILMASIN",
            "Mamul Kalitesi -FILMASIN",
            "Kutuk Kalitesi -KUTUK"
        ]
    ].copy()

# Eğitim dosyasında model isimleri varsa
elif all(
    sutun in egitim_verisi.columns
    for sutun in eski_sutunlar
):

    egitim_kombinasyonlari = egitim_verisi[
        [
            "Y_CAP_FLM_MM",
            "Y_KALITE_FLM",
            "Y_KALITE_KTK"
        ]
    ].copy()

    egitim_kombinasyonlari = egitim_kombinasyonlari.rename(
        columns={
            "Y_CAP_FLM_MM": "Filmasin Cap mm -FILMASIN",
            "Y_KALITE_FLM": "Mamul Kalitesi -FILMASIN",
            "Y_KALITE_KTK": "Kutuk Kalitesi -KUTUK"
        }
    )

else:

    st.error(
        "Diler Proje Verileri.xlsx içerisinde gerekli "
        "çap ve kalite sütunları bulunamadı."
    )

    st.stop()


# ============================================================
# GEÇERLİ KOMBİNASYONLARI TEMİZLE
# ============================================================

egitim_kombinasyonlari = (
    egitim_kombinasyonlari
    .dropna()
    .drop_duplicates()
    .copy()
)


# Karşılaştırmanın daha sağlıklı olması için
# metin alanlarını string yapıyoruz.

egitim_kombinasyonlari[
    "Mamul Kalitesi -FILMASIN"
] = (
    egitim_kombinasyonlari[
        "Mamul Kalitesi -FILMASIN"
    ]
    .astype(str)
    .str.strip()
)

egitim_kombinasyonlari[
    "Kutuk Kalitesi -KUTUK"
] = (
    egitim_kombinasyonlari[
        "Kutuk Kalitesi -KUTUK"
    ]
    .astype(str)
    .str.strip()
)


# Çapı sayısal yap
egitim_kombinasyonlari[
    "Filmasin Cap mm -FILMASIN"
] = pd.to_numeric(
    egitim_kombinasyonlari[
        "Filmasin Cap mm -FILMASIN"
    ],
    errors="coerce"
)


# Geçerli kombinasyonları set haline getir
gecerli_kombinasyonlar = set(
    zip(
        egitim_kombinasyonlari[
            "Filmasin Cap mm -FILMASIN"
        ],
        egitim_kombinasyonlari[
            "Mamul Kalitesi -FILMASIN"
        ],
        egitim_kombinasyonlari[
            "Kutuk Kalitesi -KUTUK"
        ]
    )
)


# ============================================================
# SAP ÜRÜN LİSTESİ
# ============================================================

st.subheader("SAP Ürün Listesi")

st.write(
    "SAP sisteminden alınan Excel dosyasını yükleyerek "
    "tahmin işlemini başlatın."
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

    df = pd.read_excel(dosya)


    # --------------------------------------------------------
    # GEREKLİ SÜTUNLAR
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


    # --------------------------------------------------------
    # BAŞARILI YÜKLEME
    # --------------------------------------------------------

    st.success("Excel başarıyla yüklendi.")


    # --------------------------------------------------------
    # ÜRÜN SAYISI VE BAŞLANGIÇ ZAMANI
    # --------------------------------------------------------

    col1, col2 = st.columns(2)


    with col1:

        st.metric(
            "Ürün / Parti Sayısı",
            f"{len(df):,}"
        )


    with col2:

        st.metric(
            "Eğitim Verisindeki Geçerli Kombinasyon",
            f"{len(gecerli_kombinasyonlar):,}"
        )


    st.divider()


    # ========================================================
    # KULLANICI VERİSİNİ HAZIRLA
    # ========================================================

    yeni = df[
        [
            "Filmasin Cap mm -FILMASIN",
            "Mamul Kalitesi -FILMASIN",
            "Uretilecek Paket Sayisi -FILMASIN",
            "Kutuk Kalitesi -KUTUK"
        ]
    ].copy()


    # --------------------------------------------------------
    # SAYISAL ALANLAR
    # --------------------------------------------------------

    yeni[
        "Filmasin Cap mm -FILMASIN"
    ] = pd.to_numeric(
        yeni[
            "Filmasin Cap mm -FILMASIN"
        ],
        errors="coerce"
    )


    yeni[
        "Uretilecek Paket Sayisi -FILMASIN"
    ] = pd.to_numeric(
        yeni[
            "Uretilecek Paket Sayisi -FILMASIN"
        ],
        errors="coerce"
    )


    # --------------------------------------------------------
    # KALİTE ALANLARI
    # --------------------------------------------------------

    yeni[
        "Mamul Kalitesi -FILMASIN"
    ] = (
        yeni[
            "Mamul Kalitesi -FILMASIN"
        ]
        .astype(str)
        .str.strip()
    )


    yeni[
        "Kutuk Kalitesi -KUTUK"
    ] = (
        yeni[
            "Kutuk Kalitesi -KUTUK"
        ]
        .astype(str)
        .str.strip()
    )


    # ========================================================
    # KOMBİNASYON KONTROLÜ
    # ========================================================

    yeni["TAHMIN_YAPILABILIR"] = yeni.apply(
        lambda satir: (
            satir[
                "Filmasin Cap mm -FILMASIN"
            ],
            satir[
                "Mamul Kalitesi -FILMASIN"
            ],
            satir[
                "Kutuk Kalitesi -KUTUK"
            ]
        ) in gecerli_kombinasyonlar,
        axis=1
    )


    # ========================================================
    # TAHMİN SÜTUNLARINI OLUŞTUR
    # ========================================================

    yeni[
        "TAHMINI_1_KUTUK_SURESI"
    ] = "Tahmin yapılamadı"


    yeni[
        "TAHMINI_TOPLAM_SURE"
    ] = 0.0


    # ========================================================
    # SADECE GEÇERLİ KOMBİNASYONLARI TAHMİN ET
    # ========================================================

    gecerli_index = yeni.index[
        yeni["TAHMIN_YAPILABILIR"]
    ]


    if len(gecerli_index) > 0:

        tahmin_verisi = yeni.loc[
            gecerli_index,
            [
                "Filmasin Cap mm -FILMASIN",
                "Mamul Kalitesi -FILMASIN",
                "Kutuk Kalitesi -KUTUK"
            ]
        ].copy()


        # ----------------------------------------------------
        # MODEL İÇİN ESKİ İSİMLERE ÇEVİR
        # ----------------------------------------------------

        model_input = tahmin_verisi.rename(
            columns={
                "Filmasin Cap mm -FILMASIN":
                    "Y_CAP_FLM_MM",

                "Mamul Kalitesi -FILMASIN":
                    "Y_KALITE_FLM",

                "Kutuk Kalitesi -KUTUK":
                    "Y_KALITE_KTK"
            }
        )


        # ----------------------------------------------------
        # ONE-HOT ENCODING
        # ----------------------------------------------------

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

        tahmin = model.predict(
            model_data
        )


        # ----------------------------------------------------
        # 1 KÜTÜK SÜRESİ
        # ----------------------------------------------------

        yeni.loc[
            gecerli_index,
            "TAHMINI_1_KUTUK_SURESI"
        ] = pd.Series(
            tahmin.round(2),
            index=gecerli_index
        )


        # ----------------------------------------------------
        # MİKTAR İLE ÇARP
        # ----------------------------------------------------

        toplam_tahminler = (
            tahmin
            * yeni.loc[
                gecerli_index,
                "Uretilecek Paket Sayisi -FILMASIN"
            ].values
        )


        yeni.loc[
            gecerli_index,
            "TAHMINI_TOPLAM_SURE"
        ] = pd.Series(
            toplam_tahminler.round(2),
            index=gecerli_index
        )


    # ========================================================
    # TAHMİN SAYILARI
    # ========================================================

    tahmin_yapilabilen = int(
        yeni["TAHMIN_YAPILABILIR"].sum()
    )


    tahmin_yapilamayan = (
        len(yeni)
        - tahmin_yapilabilen
    )


    # ========================================================
    # UYARI
    # ========================================================

    if tahmin_yapilamayan > 0:

        st.warning(
            f"⚠️ {tahmin_yapilamayan} ürün için tahmin "
            f"oluşturulamadı. Bu ürünlerin çap + mamul "
            f"kalitesi + kütük kalitesi kombinasyonu "
            f"Diler Proje Verileri eğitim verisinde "
            f"bulunmamaktadır. "
            f"Toplam süre yalnızca tahmin yapılabilen "
            f"{tahmin_yapilabilen} ürün üzerinden "
            f"hesaplanmıştır."
        )

    else:

        st.success(
            "Tüm ürünler için eğitim verisinde bulunan "
            "kombinasyonlara göre tahmin oluşturuldu."
        )


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
    # TOPLAM SÜRE GÖSTER
    # ========================================================

    st.metric(
        "Toplam Tahmini Üretim Süresi",
        f"{saat} sa {dakika} dk {saniye} sn"
    )


    st.divider()


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
    # TAHMİNİ BİTİŞ ZAMANI
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


    st.success(
        "🏭 Tahmini üretim bitiş zamanı: "
        + bitis_zamani.strftime(
            "%d.%m.%Y %H:%M"
        )
    )


    # ========================================================
    # TAHMİN SONUÇLARI
    # ========================================================

    st.subheader(
        "Tahmin Sonuçları"
    )


    # Kullanıcıya gösterilecek tablo
    gosterilecek_sutunlar = [
        "Filmasin Cap mm -FILMASIN",
        "Mamul Kalitesi -FILMASIN",
        "Uretilecek Paket Sayisi -FILMASIN",
        "Kutuk Kalitesi -KUTUK",
        "TAHMINI_1_KUTUK_SURESI",
        "TAHMINI_TOPLAM_SURE"
    ]


    st.dataframe(
        yeni[gosterilecek_sutunlar],
        use_container_width=True,
        hide_index=True
    )


    # ========================================================
    # EXCEL DOSYASI
    # ========================================================

    sonuc_excel = BytesIO()


    with pd.ExcelWriter(
        sonuc_excel,
        engine="openpyxl"
    ) as writer:

        yeni[gosterilecek_sutunlar].to_excel(
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

    if st.button("📧 E-posta ile Gönder"):

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

Diler Üretim Süresi Tahmin Sistemi tarafından
oluşturulan tahmin sonuçları ekte paylaşılmıştır.

Tahmin yapılabilen ürün sayısı:
{tahmin_yapilabilen}

Tahmin yapılamayan ürün sayısı:
{tahmin_yapilamayan}

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
                "Secrets ayarlarını kontrol edin."
            )

            st.write(e)