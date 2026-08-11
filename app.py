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
        padding-bottom: 50px;
    }

    h1 {
        color: #1239A6;
        font-size: 46px !important;
        font-weight: 700 !important;
    }

    h2, h3 {
        color: #172033;
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

st.write(
    "SAP üretim verileri kullanılarak ürün bazında "
    "tahmini üretim sürelerinin hesaplanması"
)

st.divider()


# ============================================================
# MODELİ YÜKLE
# ============================================================

try:

    model = joblib.load(
        "model.pkl"
    )

    model_columns = joblib.load(
        "model_columns.pkl"
    )

except Exception as e:

    st.error(
        "Model dosyaları yüklenemedi."
    )

    st.write(e)

    st.stop()


# ============================================================
# EĞİTİM VERİSİNİ YÜKLE
# ============================================================

egitim_dosyasi = (
    "Diler Proje Verileri.xlsx"
)


if not os.path.exists(
    egitim_dosyasi
):

    st.error(
        f"'{egitim_dosyasi}' dosyası "
        "uygulamanın bulunduğu klasörde bulunamadı."
    )

    st.info(
        "Diler Proje Verileri.xlsx dosyasını "
        "app.py, model.pkl ve model_columns.pkl "
        "ile aynı GitHub klasörüne yükleyin."
    )

    st.stop()


try:

    egitim_verisi = pd.read_excel(
        egitim_dosyasi
    )

except Exception as e:

    st.error(
        "Diler Proje Verileri.xlsx okunamadı."
    )

    st.write(e)

    st.stop()


# ============================================================
# EĞİTİM VERİSİ KONTROLÜ
# ============================================================
#
# Diler Proje Verileri.xlsx içerisindeki GERÇEK
# model sütunları:
#
# Y_CAP_FLM_MM
# Y_KALITE_FLM
# Y_KALITE_KTK
#
# Miktar burada aranmayacak.
# Çünkü kombinasyon kontrolü için miktara gerek yok.
#


egitim_gerekli_sutunlar = [
    "Y_CAP_FLM_MM",
    "Y_KALITE_FLM",
    "Y_KALITE_KTK"
]


eksik_egitim_sutunlari = [
    sutun
    for sutun in egitim_gerekli_sutunlar
    if sutun not in egitim_verisi.columns
]


if eksik_egitim_sutunlari:

    st.error(
        "Diler Proje Verileri.xlsx içerisinde "
        "gerekli model sütunları bulunamadı: "
        + ", ".join(
            eksik_egitim_sutunlari
        )
    )

    st.info(
        "Beklenen sütunlar: "
        "Y_CAP_FLM_MM, Y_KALITE_FLM, Y_KALITE_KTK"
    )

    st.stop()


# ============================================================
# EĞİTİM VERİSİNDEN GEÇERLİ KOMBİNASYONLARI AL
# ============================================================

egitim_kombinasyonlari = egitim_verisi[
    [
        "Y_CAP_FLM_MM",
        "Y_KALITE_FLM",
        "Y_KALITE_KTK"
    ]
].copy()


# ============================================================
# EĞİTİM VERİSİNİ TEMİZLE
# ============================================================

# Çapı sayısal yap
egitim_kombinasyonlari[
    "Y_CAP_FLM_MM"
] = pd.to_numeric(
    egitim_kombinasyonlari[
        "Y_CAP_FLM_MM"
    ],
    errors="coerce"
)


# Mamul kalitesi
egitim_kombinasyonlari[
    "Y_KALITE_FLM"
] = (
    egitim_kombinasyonlari[
        "Y_KALITE_FLM"
    ]
    .astype(str)
    .str.strip()
)


# Kütük kalitesi
egitim_kombinasyonlari[
    "Y_KALITE_KTK"
] = (
    egitim_kombinasyonlari[
        "Y_KALITE_KTK"
    ]
    .astype(str)
    .str.strip()
)


# Boş kayıtları çıkar
egitim_kombinasyonlari = (
    egitim_kombinasyonlari
    .dropna()
    .drop_duplicates()
    .copy()
)


# ============================================================
# GEÇERLİ KOMBİNASYONLAR SETİ
# ============================================================

gecerli_kombinasyonlar = set(
    zip(
        egitim_kombinasyonlari[
            "Y_CAP_FLM_MM"
        ],
        egitim_kombinasyonlari[
            "Y_KALITE_FLM"
        ],
        egitim_kombinasyonlari[
            "Y_KALITE_KTK"
        ]
    )
)


# ============================================================
# SAP ÜRÜN LİSTESİ
# ============================================================

st.subheader(
    "SAP Ürün Listesi"
)

st.write(
    "SAP sisteminden alınan Excel dosyasını "
    "yükleyerek tahmin işlemini başlatın."
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
            "Yüklenen Excel dosyası okunamadı."
        )

        st.write(e)

        st.stop()


    # --------------------------------------------------------
    # SAP EXCEL'İNDE GEREKLİ SÜTUNLAR
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

        st.info(
            "SAP Excel dosyasında olması gereken sütunlar: "
            "Filmasin Cap mm -FILMASIN, "
            "Mamul Kalitesi -FILMASIN, "
            "Uretilecek Paket Sayisi -FILMASIN, "
            "Kutuk Kalitesi -KUTUK"
        )

        st.stop()


    # --------------------------------------------------------
    # BAŞARILI YÜKLEME
    # --------------------------------------------------------

    st.success(
        "Excel başarıyla yüklendi."
    )


    # ========================================================
    # ÜRÜN SAYISI
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


    with col2:

        st.markdown(
            f"""
            <div class="card">

                <div class="card-title">
                    Eğitim Verisindeki Geçerli Kombinasyon
                </div>

                <div class="card-value">
                    {len(gecerli_kombinasyonlar):,}
                </div>

            </div>
            """,
            unsafe_allow_html=True
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
    # ÇAP
    # --------------------------------------------------------

    yeni[
        "Filmasin Cap mm -FILMASIN"
    ] = pd.to_numeric(
        yeni[
            "Filmasin Cap mm -FILMASIN"
        ],
        errors="coerce"
    )


    # --------------------------------------------------------
    # MİKTAR
    # --------------------------------------------------------

    yeni[
        "Uretilecek Paket Sayisi -FILMASIN"
    ] = pd.to_numeric(
        yeni[
            "Uretilecek Paket Sayisi -FILMASIN"
        ],
        errors="coerce"
    )


    # --------------------------------------------------------
    # MAMUL KALİTESİ
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


    # --------------------------------------------------------
    # KÜTÜK KALİTESİ
    # --------------------------------------------------------

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
    #
    # Burada yalnızca:
    #
    # Çap
    # +
    # Mamul Kalitesi
    # +
    # Kütük Kalitesi
    #
    # üçlüsünün eğitim verisinde bulunup bulunmadığı
    # kontrol ediliyor.
    #
    # KTKID KESİNLİKLE KULLANILMIYOR.
    #


    yeni[
        "TAHMIN_YAPILABILIR"
    ] = yeni.apply(
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
    # TAHMİN SÜTUNLARINI BAŞLAT
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
        yeni[
            "TAHMIN_YAPILABILIR"
        ]
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

        try:

            tahmin = model.predict(
                model_data
            )

        except Exception as e:

            st.error(
                "Model tahmin sırasında hata oluştu."
            )

            st.write(e)

            st.stop()


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

        miktarlar = yeni.loc[
            gecerli_index,
            "Uretilecek Paket Sayisi -FILMASIN"
        ].fillna(0).values


        toplam_tahminler = (
            tahmin
            * miktarlar
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
        yeni[
            "TAHMIN_YAPILABILIR"
        ].sum()
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
            f"bulunmamaktadır. Toplam süre yalnızca "
            f"tahmin yapılabilen {tahmin_yapilabilen} "
            f"ürün üzerinden hesaplanmıştır."
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
    # TOPLAM SÜRE KARTI
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


    gosterilecek_sutunlar = [
        "Filmasin Cap mm -FILMASIN",
        "Mamul Kalitesi -FILMASIN",
        "Uretilecek Paket Sayisi -FILMASIN",
        "Kutuk Kalitesi -KUTUK",
        "TAHMINI_1_KUTUK_SURESI",
        "TAHMINI_TOPLAM_SURE"
    ]


    st.dataframe(
        yeni[
            gosterilecek_sutunlar
        ],
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

        yeni[
            gosterilecek_sutunlar
        ].to_excel(
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