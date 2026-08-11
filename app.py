#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Aug 11 08:26:39 2026

@author: selinaysancak
"""
import streamlit as st
import pandas as pd
import joblib
from io import BytesIO

# -------------------------------------------------
# SAYFA AYARLARI
# -------------------------------------------------

st.set_page_config(
    page_title="Diler | Üretim Süresi Tahmin Sistemi",
    page_icon="Diler_Logo.png",
    layout="wide"
)

# -------------------------------------------------
# TASARIM
# -------------------------------------------------

st.markdown("""
<style>

    /* Genel sayfa */
    .stApp {
        background-color: #ffffff;
    }

    .block-container {
        max-width: 1200px;
        padding-top: 2rem;
        padding-bottom: 4rem;
    }

    /* Üst logo alanı */
    .logo-area {
        display: flex;
        align-items: center;
        padding: 10px 0 25px 0;
        border-bottom: 1px solid #eeeeee;
        margin-bottom: 45px;
    }

    /* Ana başlık */
    .main-title {
        color: #1239a6;
        font-size: 48px;
        font-weight: 700;
        line-height: 1.1;
        margin-bottom: 10px;
    }

    .subtitle {
        color: #5f6470;
        font-size: 18px;
        margin-bottom: 40px;
    }

    /* Bölüm başlıkları */
    .section-title {
        color: #172033;
        font-size: 28px;
        font-weight: 700;
        margin-top: 30px;
        margin-bottom: 8px;
    }

    .section-text {
        color: #6b7280;
        font-size: 16px;
        margin-bottom: 20px;
    }

    /* Bilgi kartları */
    .metric-card {
        background: #f5f7fc;
        border-radius: 16px;
        padding: 25px;
        border: 1px solid #e5e8f0;
        height: 100%;
    }

    .metric-title {
        color: #6b7280;
        font-size: 15px;
        margin-bottom: 8px;
    }

    .metric-value {
        color: #1239a6;
        font-size: 32px;
        font-weight: 700;
    }

    /* Sonuç kutusu */
    .result-box {
        background: #1239a6;
        border-radius: 18px;
        padding: 28px;
        color: white;
        margin-top: 25px;
        margin-bottom: 30px;
    }

    .result-title {
        font-size: 16px;
        opacity: 0.85;
    }

    .result-value {
        font-size: 34px;
        font-weight: 700;
        margin-top: 5px;
    }

    /* Streamlit buton */
    .stDownloadButton > button {
        background-color: #1239a6 !important;
        color: white !important;
        border: none !important;
        border-radius: 30px !important;
        padding: 12px 28px !important;
        font-weight: 600 !important;
        font-size: 16px !important;
    }

    .stDownloadButton > button:hover {
        background-color: #0d2d85 !important;
    }

    /* Dosya yükleme alanı */
    [data-testid="stFileUploader"] {
        background-color: #f5f7fc;
        border: 2px dashed #cbd3e6;
        border-radius: 18px;
        padding: 20px;
    }

</style>
""", unsafe_allow_html=True)


# -------------------------------------------------
# LOGO
# -------------------------------------------------

st.image("Diler_Logo_duzeltilmis.png", width=300)


# -------------------------------------------------
# BAŞLIK
# -------------------------------------------------

st.markdown(
    '<div class="main-title">Üretim Süresi<br>Tahmin Sistemi</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="subtitle">'
    'SAP üretim verileri kullanılarak ürün bazında tahmini üretim sürelerini hesaplayın.'
    '</div>',
    unsafe_allow_html=True
)


# -------------------------------------------------
# MODEL
# -------------------------------------------------

model = joblib.load("model.pkl")
model_columns = joblib.load("model_columns.pkl")


# -------------------------------------------------
# EXCEL YÜKLEME
# -------------------------------------------------

st.markdown(
    '<div class="section-title">SAP Ürün Listesi</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="section-text">'
    'SAP sisteminden alınan Excel dosyasını yükleyerek tahmin işlemini başlatın.'
    '</div>',
    unsafe_allow_html=True
)

dosya = st.file_uploader(
    "Excel dosyasını seçin",
    type=["xlsx"],
    label_visibility="collapsed"
)


# -------------------------------------------------
# DOSYA YÜKLENDİ
# -------------------------------------------------

if dosya is not None:

    df = pd.read_excel(dosya)

    gerekli_sutunlar = [
        "KTKID",
        "Y_CAP_FLM_MM",
        "Y_KALITE_FLM",
        "Y_KALITE_KTK"
    ]

    eksik = [
        sutun for sutun in gerekli_sutunlar
        if sutun not in df.columns
    ]

    if eksik:

        st.error(
            "Excel dosyasında gerekli sütunlar bulunamadı: "
            + ", ".join(eksik)
        )

    else:

        st.success("Excel başarıyla yüklendi.")

        # -----------------------------------------
        # BİLGİ KARTLARI
        # -----------------------------------------

        col1, col2 = st.columns(2)

        with col1:

            st.markdown(
                f"""
                <div class="metric-card">
                    <div class="metric-title">Ürün Sayısı</div>
                    <div class="metric-value">{len(df):,}</div>
                </div>
                """,
                unsafe_allow_html=True
            )

        # -----------------------------------------
        # MODEL VERİSİ
        # -----------------------------------------

        yeni = df[
            [
                "KTKID",
                "Y_CAP_FLM_MM",
                "Y_KALITE_FLM",
                "Y_KALITE_KTK"
            ]
        ].copy()

        model_data = pd.get_dummies(
            yeni[
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

        # -----------------------------------------
        # TAHMİN
        # -----------------------------------------

        tahmin = model.predict(model_data)

        yeni["TAHMINI_URUN_SURESI"] = tahmin.round(2)

        toplam_sure = tahmin.sum()

        saat = int(toplam_sure // 3600)
        dakika = int((toplam_sure % 3600) // 60)
        saniye = int(toplam_sure % 60)

        with col2:

            st.markdown(
                f"""
                <div class="metric-card">
                    <div class="metric-title">
                        Toplam Tahmini Süre
                    </div>
                    <div class="metric-value">
                        {saat} sa {dakika} dk
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )

        # -----------------------------------------
        # SONUÇ
        # -----------------------------------------

        st.markdown(
            '<div class="section-title">Tahmin Sonuçları</div>',
            unsafe_allow_html=True
        )

        st.dataframe(
            yeni,
            use_container_width=True,
            hide_index=True
        )

        # -----------------------------------------
        # TOPLAM SÜRE
        # -----------------------------------------

        st.markdown(
            f"""
            <div class="result-box">
                <div class="result-title">
                    TOPLAM TAHMİNİ ÜRETİM SÜRESİ
                </div>
                <div class="result-value">
                    {saat} saat {dakika} dakika {saniye} saniye
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        # -----------------------------------------
        # EXCEL OLUŞTUR
        # -----------------------------------------

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

        st.download_button(
            label="Tahmin Sonuçlarını Excel Olarak İndir",
            data=sonuc_excel,
            file_name="Diler_Tahmin_Sonuclari.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )