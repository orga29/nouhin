import tempfile
from pathlib import Path

import streamlit as st
from nouhin import prepare            # ← 追加したラッパーを呼ぶ

st.title("📦 納品数シート自動化ツール")

delivery_date = st.date_input("納品日")
upload        = st.file_uploader("在庫集計表 (.xlsm)", type=["xlsm"])

if st.button("実行") and upload:
    # アップロード → 一時ファイルへ保存
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsm") as tf:
        tf.write(upload.getbuffer())
        src_path = Path(tf.name)

    try:
        out_path = prepare(src_path, delivery_date.isoformat())
    except Exception as e:
        st.error("🚨 コア処理でエラーが発生しました")
        st.exception(e)
        st.stop()

    dl_name = f"{delivery_date.strftime('%Y%m%d')}納品.xlsm"
    with open(out_path, "rb") as f:
        st.success("✅ 完了しました！")
        st.download_button(
            "⬇️ ダウンロード", f, file_name=dl_name,
            mime="application/vnd.ms-excel.sheet.macroEnabled.12"
        )
