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

st.set_page_config(
    page_title="Diler | Üretim Süresi Tahmin Sistemi",
    page_icon="Diler_Logo_duzeltilmis.png",
    layout="wide"
)

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

st.image("Diler_Logo_duzeltilmis.png", width=300)

st.title("Üretim Süresi Tahmin Sistemi")

st.markdown(
    '<div class="subtitle">'
    'SAP üretim verileri kullanılarak ürün bazında tahmini üretim sürelerinin hesaplanması'
    '</div>',
    unsafe_allow_html=True
)

st.divider()

model = joblib.load("model.pkl")
model_columns = joblib.load("model_columns.pkl")

st.subheader("SAP Ürün Listesi")

st.write(
    "SAP sisteminden alınan Excel dosyasını yükleyerek tahmin işlemini başlatın."
)

dosya = st.file_uploader(
    "Excel dosyasını seçin",
    type=["xlsx"],
    label_visibility="collapsed"
)

if dosya is not None:

    df = pd.read_excel(dosya)

    gerekli_sutunlar = [
        "KTKID",
        "Y_CAP_FLM_MM",
        "Y_KALITE_FLM",
        "Y_KALITE_KTK",
        "Miktar"
    ]

    eksik = [
        sutun for sutun in gerekli_sutunlar
        if sutun not in df.columns
    ]

    if eksik:

        st.error(
            "Excel dosyasında şu sütunlar eksik: "
            + ", ".join(eksik)
        )

    else:

        st.success("Excel başarıyla yüklendi.")

        col1, col2 = st.columns(2)

        with col1:
            st.markdown(
                f"""
                <div class="card">
                    <div class="card-title">Ürün / Parti Sayısı</div>
                    <div class="card-value">{len(df):,}</div>
                </div>
                """,
                unsafe_allow_html=True
            )

        yeni = df[
            [
                "KTKID",
                "Y_CAP_FLM_MM",
                "Y_KALITE_FLM",
                "Y_KALITE_KTK",
                "Miktar"
            ]
        ].copy()

        # Miktarı sayıya çevir
        yeni["Miktar"] = pd.to_numeric(
            yeni["Miktar"],
            errors="coerce"
        )

        # Tahmin için model verisi
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

        # 1 kütük için tahmin
        tahmin = model.predict(model_data)

        yeni["TAHMINI_1_KUTUK_SURESI"] = tahmin.round(2)

        # Miktar ile çarp
        yeni["TAHMINI_TOPLAM_SURE"] = (
            tahmin * yeni["Miktar"]
        ).round(2)

        # Toplam süre
        toplam_sure = yeni["TAHMINI_TOPLAM_SURE"].sum()

        saat = int(toplam_sure // 3600)
        dakika = int((toplam_sure % 3600) // 60)
        saniye = int(toplam_sure % 60)

        with col2:
            st.markdown(
                f"""
                <div class="card">
                    <div class="card-title">Toplam Tahmini Süre</div>
                    <div class="card-value">
                        {saat} sa {dakika} dk
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )

        st.divider()

        st.subheader("Tahmin Sonuçları")

        st.dataframe(
            yeni,
            use_container_width=True,
            hide_index=True
        )

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

        # Excel oluştur
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