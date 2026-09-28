import streamlit as st

from ais_extractor import build_ais_workbook, read_ais_json

st.set_page_config(page_title="AIS JSON Extractor", page_icon="🧾", layout="centered")

st.markdown(
    """
<style>
.main { max-width: 760px; }
.stApp { background: #f8fafc; }
.title-box { background: linear-gradient(135deg, #1a56db, #1e429f); color: white;
             border-radius: 14px; padding: 28px 32px; margin-bottom: 24px; text-align: center; }
.title-box h1 { font-size: 28px; margin: 0; font-weight: 700; }
.title-box p { margin: 6px 0 0; opacity: .85; font-size: 14px; }
</style>
""",
    unsafe_allow_html=True,
)

st.markdown(
    """
<div class="title-box">
  <h1>🧾 AIS JSON Extractor</h1>
  <p>Convert Income Tax portal AIS reported data into a reviewable Excel workbook.</p>
</div>
""",
    unsafe_allow_html=True,
)
st.info(
    "Upload the readable JSON downloaded from the AIS portal. Encrypted or non-JSON "
    "downloads must first be opened/exported using the official AIS Utility."
)
st.markdown(
    "The generated workbook includes an **Index** and a worksheet for every "
    "reported-data category discovered in the file."
)

ais_file = st.file_uploader("Upload AIS JSON file", type=["json"])
if ais_file:
    st.markdown(f"**File:** `{ais_file.name}` · `{ais_file.size / 1024:.0f} KB`")
    if st.button("Extract reported data", type="primary", use_container_width=True):
        try:
            with st.spinner("Reading AIS data and creating Excel workbook…"):
                output, row_counts = build_ais_workbook(read_ais_json(ais_file.getvalue()))
            st.success(
                f"Extracted {sum(row_counts.values())} reported rows across "
                f"{len(row_counts)} categories."
            )
            st.download_button(
                "⬇️ Download AIS reported data (.xlsx)",
                output,
                "AIS_REPORTED_DATA.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        except ValueError as error:
            st.error(str(error))
        except Exception as error:
            st.error(f"Could not extract the AIS file: {error}")
