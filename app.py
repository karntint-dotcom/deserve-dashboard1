import streamlit as st
import pandas as pd
import io
import traceback
from engine import build_dashboard

st.set_page_config(
    page_title="Deserve Sales Dashboard Builder",
    page_icon="📊",
    layout="centered",
)

# ── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { max-width: 720px; margin: auto; }
    .big-title { font-size: 2rem; font-weight: 700; color: #1F4E79; margin-bottom: 0; }
    .sub-title { font-size: 1rem; color: #595959; margin-bottom: 1.5rem; }
    .step-box {
        background: #F0F4FF; border-left: 4px solid #2E75B6;
        padding: 12px 16px; border-radius: 6px; margin-bottom: 12px;
    }
    .step-num { font-weight: 700; color: #2E75B6; }
    .info-chip {
        display: inline-block; background: #E2EFDA; color: #375623;
        border-radius: 12px; padding: 2px 10px; font-size: 0.8rem;
        margin: 2px; font-weight: 600;
    }
    .warn-chip {
        display: inline-block; background: #FCE4D6; color: #843C0C;
        border-radius: 12px; padding: 2px 10px; font-size: 0.8rem;
        margin: 2px; font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# ── Header ───────────────────────────────────────────────────────────────────
st.markdown('<p class="big-title">📊 Deserve Dashboard Builder</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">อัปโหลดไฟล์ยอดขาย → ได้ Excel Dashboard พร้อมใช้ทันที</p>', unsafe_allow_html=True)

# ── What you get ─────────────────────────────────────────────────────────────
with st.expander("📋 Dashboard ที่ได้มีอะไรบ้าง?", expanded=False):
    sheets = [
        ("1", "ภาพรวมยอดขาย", "KPI รายเดือน ยอดรวม จำนวนร้าน"),
        ("2", "Ranking ร้านค้า", "จัดอันดับตามชื่อร้าน + สถานะ trend"),
        ("3", "Ranking รหัสลูกค้า", "จัดอันดับตามรหัสลูกค้า พร้อมชื่อร้าน"),
        ("4", "ยอดขายฟาร์ม", "แยกเฉพาะลูกค้าที่รหัสขึ้นต้นด้วย 'ฟาร์ม'"),
        ("5", "เป้าหมาย PLC & Recco", "ติดตามยอด Pet Lover Centre vs เรคโค"),
        ("6", "ภาพรวมสินค้า", "แยกตามหมวดหมู่สินค้า"),
        ("7", "Ranking สินค้า", "จัดอันดับสินค้า (กลุ่ม 6 ตัวแรก)"),
        ("8", "สินค้า × Top10 ร้าน", "สินค้าที่ขายดีในแต่ละร้าน Top 10"),
        ("9", "สินค้า × Top10 รหัส", "สินค้าที่ขายดีในแต่ละรหัส Top 10"),
        ("10", "Forecast ร้านค้า", "🌟 STAR / 📈 GROWTH / 📉 DECLINING + forecast มิ.ย.–ธ.ค."),
        ("11", "Forecast สินค้า × ร้าน", "สินค้าไหนขายดีร้านไหน / หยุดซื้อ / กำลังโต"),
    ]
    for num, name, desc in sheets:
        st.markdown(f"**Sheet {num}** — {name}: {desc}")

# ── File format guide ─────────────────────────────────────────────────────────
with st.expander("📁 รูปแบบไฟล์ที่รองรับ", expanded=False):
    st.markdown("""
ไฟล์ Excel **.xlsx** ที่มี **3 sheets** ในชื่อนี้:

| Sheet | ชื่อ | คำอธิบาย |
|-------|------|-----------|
| 1 | `ข้อมูลเดือน 1-5` | ข้อมูลยอดขาย (header row 2) |
| 2 | `เงื่อนไข1` | Mapping ชื่อลูกค้า → ชื่อร้าน |
| 3 | `เงื่อนไข2` | Mapping รหัสลูกค้า (สาขาเดียวกัน) |

**คอลัมน์หลักที่ต้องมีใน Sheet 1:**
`ชื่อลูกค้า`, `รหัสลูกค้า`, `วันที่ทำรายการ`, `รหัสสินค้า`, `ชื่อสินค้า`, `จำนวน`, `ราคารวม`, `สถานะรายการ`, `หมวดหมู่`
""")

# ── Target settings ───────────────────────────────────────────────────────────
st.markdown("### ⚙️ ตั้งค่าเป้าหมายยอดขาย")
col1, col2 = st.columns(2)
with col1:
    plc_target = st.number_input(
        "🎯 Pet Lover Centre (บาท)",
        value=1_250_000, step=50_000, format="%d",
        help="เป้าหมายยอดขาย Pet Lover Centre ถึง 30 มิ.ย."
    )
    plc_deadline = st.text_input("กำหนดเวลา", value="30 มิถุนายน 2569")
with col2:
    recco_target = st.number_input(
        "🎯 บริษัท เรคโค เพ็ท จำกัด (บาท)",
        value=1_000_000, step=50_000, format="%d",
        help="เป้าหมายยอดขาย เรคโค เพ็ท ถึง 31 ธ.ค."
    )
    recco_deadline = st.text_input("กำหนดเวลา ", value="31 ธันวาคม 2569")

# ── Upload ────────────────────────────────────────────────────────────────────
st.markdown("### 📤 อัปโหลดไฟล์ข้อมูล")

uploaded = st.file_uploader(
    "เลือกไฟล์ .xlsx",
    type=["xlsx"],
    help="ไฟล์ยอดขายที่มี sheet ข้อมูลเดือน 1-5, เงื่อนไข1, เงื่อนไข2",
)

if uploaded:
    st.success(f"✅ ได้รับไฟล์: **{uploaded.name}** ({uploaded.size/1024:.1f} KB)")

    # Quick preview
    try:
        df_prev = pd.read_excel(uploaded, sheet_name='ข้อมูลเดือน 1-5', header=1, nrows=3)
        uploaded.seek(0)
        with st.expander("👀 Preview ข้อมูล (3 แถวแรก)"):
            st.dataframe(df_prev[['ชื่อลูกค้า','รหัสลูกค้า','วันที่ทำรายการ',
                                   'รหัสสินค้า','ชื่อสินค้า','ราคารวม']].head(3),
                         use_container_width=True)
        uploaded.seek(0)
    except Exception:
        uploaded.seek(0)

    # Build button
    st.markdown("---")
    if st.button("🚀 สร้าง Dashboard", type="primary", use_container_width=True):
        progress = st.progress(0, text="เริ่มประมวลผล...")
        status   = st.empty()

        try:
            file_bytes = uploaded.read()
            steps = [
                (10,  "📥 โหลดข้อมูลและเงื่อนไข..."),
                (25,  "🏪 วิเคราะห์ร้านค้า..."),
                (40,  "📦 วิเคราะห์สินค้า..."),
                (55,  "📊 สร้าง Ranking sheets..."),
                (70,  "🎯 คำนวณเป้าหมาย..."),
                (85,  "🔮 Forecast มิ.ย.–ธ.ค...."),
                (95,  "🔗 ผูกสูตร Excel..."),
                (100, "✅ เสร็จสมบูรณ์!"),
            ]

            def update(pct, msg):
                progress.progress(pct, text=msg)
                status.markdown(f"_{msg}_")

            output_bytes = build_dashboard(
                file_bytes=file_bytes,
                plc_target=plc_target,
                plc_deadline=plc_deadline,
                recco_target=recco_target,
                recco_deadline=recco_deadline,
                progress_cb=update,
            )

            progress.progress(100, text="✅ เสร็จสมบูรณ์!")
            status.empty()

            # Show summary metrics
            st.balloons()
            st.success("🎉 สร้าง Dashboard สำเร็จ!")

            # Download button
            fname = uploaded.name.replace('.xlsx','') + "_Dashboard.xlsx"
            st.download_button(
                label="⬇️ ดาวน์โหลด Dashboard (.xlsx)",
                data=output_bytes,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )

        except Exception as e:
            progress.empty()
            status.empty()
            st.error(f"❌ เกิดข้อผิดพลาด: {e}")
            with st.expander("🔍 รายละเอียด error"):
                st.code(traceback.format_exc())

else:
    st.info("👆 อัปโหลดไฟล์ข้อมูลเพื่อเริ่มต้น")

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    "<p style='text-align:center; color:#AAAAAA; font-size:0.8rem;'>"
    "Deserve Dashboard Builder · สร้างด้วย Python + Streamlit"
    "</p>",
    unsafe_allow_html=True,
)
