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
import textwrap

from io import BytesIO
from pathlib import Path
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
# DOSYA YOLLARI
# ============================================================

MODEL_FILE = "model.pkl"
MODEL_COLUMNS_FILE = "model_columns.pkl"
TRAINING_FILE = "Diler Proje Verileri.xlsx"
LOGO_FILE = "Diler_Logo_duzeltilmis.png"


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
        margin-bottom: 8px;
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
        margin-bottom: 8px;
    }

    .total-value {
        font-size: 30px;
        font-weight: 700;
    }

    .info-box {
        background-color: #F4F6FB;
        border: 1px solid #E1E5EF;
        border-radius: 15px;
        padding: 20px;
        margin-top: 20px;
        margin-bottom: 20px;
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

if Path(LOGO_FILE).exists():
    st.image(LOGO_FILE, width=300)

st.title("Üretim Süresi Tahmin Sistemi")

st.markdown(
    '<div class="subtitle">'
    'SAP üretim verileri kullanılarak ürün bazında tahmini üretim sürelerinin hesaplanması'
    '</div>',
    unsafe_allow_html=True
)

st.divider()


# ============================================================
# MODEL DOSYALARINI YÜKLE
# ============================================================

if not Path(MODEL_FILE).exists():
    st.error(
        "model.pkl dosyası bulunamadı. "
        "Dosyanın GitHub projesinde bulunduğundan emin olun."
    )
    st.stop()


if not Path(MODEL_COLUMNS_FILE).exists():
    st.error(
        "model_columns.pkl dosyası bulunamadı. "
        "Dosyanın GitHub projesinde bulunduğundan emin olun."
    )
    st.stop()


model = joblib.load(MODEL_FILE)
model_columns = joblib.load(MODEL_COLUMNS_FILE)


# ============================================================
# EĞİTİM VERİSİNİ YÜKLE
# ============================================================

if not Path(TRAINING_FILE).exists():

    st.error(
        "Diler Proje Verileri.xlsx dosyası bulunamadı. "
        "Bu dosyayı GitHub projesine app.py ile aynı klasöre yükleyin."
    )

    st.stop()


try:

    egitim_df = pd.read_excel(
        TRAINING_FILE
    )

except Exception as e:

    st.error(
        "Diler Proje Verileri.xlsx okunamadı."
    )

    st.write(e)

    st.stop()


# ============================================================
# EĞİTİM VERİSİNDE GEREKLİ SÜTUNLAR
# ============================================================

egitim_sutunlari = [
    "Y_CAP_FLM_MM",
    "Y_KALITE_FLM",
    "Y_KALITE_KTK"
]


eksik_egitim = [
    sutun
    for sutun in egitim_sutunlari
    if sutun not in egitim_df.columns
]


if eksik_egitim:

    st.error(
        "Diler Proje Verileri.xlsx içerisinde gerekli "
        "model sütunları bulunamadı: "
        + ", ".join(eksik_egitim)
    )

    st.stop()


# ============================================================
# KOMBINASYONLARI NORMALİZE ET
# ============================================================

def normalize_value(value):

    if pd.isna(value):
        return ""

    if isinstance(value, float):

        if value.is_integer():
            return str(int(value))

        return str(value).strip()

    return str(value).strip().upper()


egitim_df["_CAP_KEY"] = (
    egitim_df["Y_CAP_FLM_MM"]
    .apply(normalize_value)
)

egitim_df["_FLM_KEY"] = (
    egitim_df["Y_KALITE_FLM"]
    .apply(normalize_value)
)

egitim_df["_KTK_KEY"] = (
    egitim_df["Y_KALITE_KTK"]
    .apply(normalize_value)
)


# ============================================================
# EĞİTİM VERİSİNDE BULUNAN GEÇERLİ KOMBİNASYONLAR
# ============================================================

gecerli_kombinasyonlar = set(
    zip(
        egitim_df["_CAP_KEY"],
        egitim_df["_FLM_KEY"],
        egitim_df["_KTK_KEY"]
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


    # ========================================================
    # KULLANICI EXCEL'İNDE OLMASI GEREKEN SÜTUNLAR
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

        st.info(
            "Excel dosyanızdaki sütun adları tam olarak şu şekilde olmalıdır:"
        )

        for sutun in gerekli_sutunlar:
            st.write(f"• {sutun}")

        st.stop()


    st.success(
        "Excel başarıyla yüklendi."
    )


    # ========================================================
    # SONUÇ TABLOSU
    # ========================================================

    yeni = df[
        gerekli_sutunlar
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
    # KOMBİNASYON ANAHTARLARI
    # ========================================================

    yeni["_CAP_KEY"] = (
        yeni[
            "Filmasin Cap mm -FILMASIN"
        ]
        .apply(normalize_value)
    )

    yeni["_FLM_KEY"] = (
        yeni[
            "Mamul Kalitesi -FILMASIN"
        ]
        .apply(normalize_value)
    )

    yeni["_KTK_KEY"] = (
        yeni[
            "Kutuk Kalitesi -KUTUK"
        ]
        .apply(normalize_value)
    )


    yeni["_KOMB_KEY"] = list(
        zip(
            yeni["_CAP_KEY"],
            yeni["_FLM_KEY"],
            yeni["_KTK_KEY"]
        )
    )


    # ========================================================
    # KOMBINASYON KONTROLÜ
    # ========================================================

    yeni[
        "TAHMIN_YAPILABILIR"
    ] = yeni[
        "_KOMB_KEY"
    ].isin(
        gecerli_kombinasyonlar
    )


    # ========================================================
    # TAHMİN SÜTUNLARINI BAŞLANGIÇTA 0 YAP
    # ========================================================

    yeni[
        "TAHMINI_1_KUTUK_SURESI"
    ] = 0.0


    yeni[
        "TAHMINI_TOPLAM_SURE"
    ] = 0.0


    # ========================================================
    # SADECE GEÇERLİ KOMBINASYONLAR TAHMİN EDİLECEK
    # ========================================================

    gecerli_index = yeni.index[
        yeni[
            "TAHMIN_YAPILABILIR"
        ]
    ]


    if len(gecerli_index) > 0:

        # -----------------------------------------------
        # MODEL İÇİN VERİ HAZIRLA
        # -----------------------------------------------

        model_input = yeni.loc[
            gecerli_index,
            [
                "Filmasin Cap mm -FILMASIN",
                "Mamul Kalitesi -FILMASIN",
                "Kutuk Kalitesi -KUTUK"
            ]
        ].copy()


        # -----------------------------------------------
        # MODELİN ESKİ SÜTUN İSİMLERİ
        # -----------------------------------------------

        model_input = model_input.rename(
            columns={
                "Filmasin Cap mm -FILMASIN":
                    "Y_CAP_FLM_MM",

                "Mamul Kalitesi -FILMASIN":
                    "Y_KALITE_FLM",

                "Kutuk Kalitesi -KUTUK":
                    "Y_KALITE_KTK"
            }
        )


        # -----------------------------------------------
        # ONE-HOT ENCODING
        # -----------------------------------------------

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


        # -----------------------------------------------
        # MODEL SÜTUNLARIYLA EŞLEŞTİR
        # -----------------------------------------------

        model_data = model_data.reindex(
            columns=model_columns,
            fill_value=0
        )


        # -----------------------------------------------
        # MODEL TAHMİNİ
        # -----------------------------------------------

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


        tahmin = tahmin.flatten()


        # -----------------------------------------------
        # SONUÇLARI TEK TEK YAZ
        # -----------------------------------------------

        for i, index in enumerate(
            gecerli_index
        ):

            tahmin_suresi = round(
                float(tahmin[i]),
                2
            )


            yeni.at[
                index,
                "TAHMINI_1_KUTUK_SURESI"
            ] = tahmin_suresi


            miktar = yeni.at[
                index,
                "Uretilecek Paket Sayisi -FILMASIN"
            ]


            if pd.isna(miktar):

                toplam = 0.0

            else:

                toplam = round(
                    tahmin_suresi
                    * float(miktar),
                    2
                )


            yeni.at[
                index,
                "TAHMINI_TOPLAM_SURE"
            ] = toplam


    # ========================================================
    # TAHMİN DURUMU
    # ========================================================

    yeni[
        "Tahmin Durumu"
    ] = yeni[
        "TAHMIN_YAPILABILIR"
    ].map(
        {
            True: "Tahmin Yapıldı",
            False: "Tahmin Yapılamadı"
        }
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
    # İSTATİSTİKLER
    # ========================================================

    toplam_urun = len(
        yeni
    )


    tahmin_yapilan = int(
        yeni[
            "TAHMIN_YAPILABILIR"
        ].sum()
    )


    tahmin_yapilamayan = (
        toplam_urun
        - tahmin_yapilan
    )


    # ========================================================
    # ÜST BİLGİ KARTLARI
    # ========================================================

    col1, col2, col3 = st.columns(3)


with col1:
    st.markdown(
        textwrap.dedent(f"""
        <div class="card">
            <div class="card-title">Toplam Ürün / Parti</div>
            <div class="card-value">{toplam_urun:,}</div>
        </div>
        """),
        unsafe_allow_html=True
    )


with col2:
    st.markdown(
        textwrap.dedent(f"""
        <div class="card">
            <div class="card-title">Tahmin Yapılabilen</div>
            <div class="card-value">{tahmin_yapilan:,}</div>
        </div>
        """),
        unsafe_allow_html=True
    )


with col3:
    st.markdown(
        textwrap.dedent(f"""
        <div class="card">
            <div class="card-title">Tahmin Yapılamayan</div>
            <div class="card-value">{tahmin_yapilamayan:,}</div>
        </div>
        """),
        unsafe_allow_html=True
    )


    # ========================================================
    # UYARI
    # ========================================================

    if tahmin_yapilamayan > 0:

        st.warning(
            f"{tahmin_yapilamayan} ürün için tahmin yapılamadı. "
            "Bu ürünlerin çap + mamul kalitesi + kütük kalitesi "
            "kombinasyonu Diler Proje Verileri eğitim verisinde "
            "bulunmamaktadır. Bu ürünlerin tahmini süresi 0 olarak "
            "kabul edilmiş ve toplam süreye dahil edilmemiştir."
        )


    # ========================================================
    # ÜRETİM BAŞLANGIÇ BİLGİLERİ
    # ========================================================

    st.divider()

    st.subheader(
        "Üretim Başlangıç Bilgileri"
    )


    tarih_col, saat_col = st.columns(2)


    with tarih_col:

        baslangic_tarihi = st.date_input(
            "Başlangıç Tarihi",
            value=date.today(),
            format="DD.MM.YYYY"
        )


    with saat_col:

        baslangic_saati = st.time_input(
            "Başlangıç Saati",
            value=time(8, 0)
        )


    # ========================================================
    # BİTİŞ ZAMANI
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
    # TOPLAM SÜRE
    # ========================================================

    st.markdown(
    textwrap.dedent(f"""
    <div class="total-box">
        <div class="total-title">
            TOPLAM TAHMİNİ ÜRETİM SÜRESİ
        </div>
        <div class="total-value">
            {saat} saat {dakika} dakika {saniye} saniye
        </div>
    </div>
    """),
    unsafe_allow_html=True
)


    # ========================================================
    # TAHMİNİ BİTİŞ ZAMANI
    # ========================================================

    st.markdown(
    textwrap.dedent(f"""
    <div class="total-box">
        <div class="total-title">
            TAHMİNİ ÜRETİM BİTİŞ ZAMANI
        </div>
        <div class="total-value">
            {bitis_metni}
        </div>
    </div>
    """),
    unsafe_allow_html=True
)


    # ========================================================
    # SONUÇ TABLOSU
    # ========================================================

    st.divider()

    st.subheader(
        "Tahmin Sonuçları"
    )


    gosterilecek_sutunlar = [
        "Filmasin Cap mm -FILMASIN",
        "Mamul Kalitesi -FILMASIN",
        "Uretilecek Paket Sayisi -FILMASIN",
        "Kutuk Kalitesi -KUTUK",
        "TAHMINI_1_KUTUK_SURESI",
        "TAHMINI_TOPLAM_SURE",
        "Tahmin Durumu"
    ]


    st.dataframe(
        yeni[
            gosterilecek_sutunlar
        ],
        use_container_width=True,
        hide_index=True
    )


    # ========================================================
    # EXCEL OLUŞTUR
    # ========================================================

    sonuc_df = yeni[
        gosterilecek_sutunlar
    ].copy()


    sonuc_excel = BytesIO()


    with pd.ExcelWriter(
        sonuc_excel,
        engine="openpyxl"
    ) as writer:

        sonuc_df.to_excel(
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


            mesaj[
                "Subject"
            ] = (
                "Diler Üretim Süresi "
                "Tahmin Sonuçları"
            )


            mesaj[
                "From"
            ] = email_address


            mesaj[
                "To"
            ] = to_email


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

Tahmin yapılamayan ürünlerin kombinasyonları
eğitim verisinde bulunmadığı için bu ürünlerin
üretim süresi 0 kabul edilmiştir.

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
                "Secrets ayarlarınızı kontrol edin."
            )

            st.write(e)


    # ========================================================
    # GEÇİCİ KOLONLARI TEMİZLE
    # ========================================================

    # Bu işlem sadece bellekteki dataframe için yapılır.
    # Kullanıcıya gösterilen tablo zaten temizdir.

    yeni.drop(
        columns=[
            "_CAP_KEY",
            "_FLM_KEY",
            "_KTK_KEY",
            "_KOMB_KEY",
            "TAHMIN_YAPILABILIR"
        ],
        inplace=True,
        errors="ignore"
    )