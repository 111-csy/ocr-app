import streamlit as st
import pandas as pd
import re
import io
import json
import base64
from openai import OpenAI

st.set_page_config(page_title="盘点智能识别系统", layout="wide")
st.title("📝 盘点表智能识别与业务逻辑核算系统")

# ==========================================
# ⚠️ 注意：请在这里替换你的密钥！！！
# ==========================================
client = OpenAI(
    api_key="5a238255e26546b78053e509658d247f.pbr3Zg9LPRsy0Ywx", 
    base_url="https://open.bigmodel.cn/api/paas/v4"
)
# ==========================================

try:
    db = pd.read_excel("data.xlsx", dtype=str)
    db = db.fillna("")
except Exception as e:
    db = pd.DataFrame()

if 'ocr_data' not in st.session_state:
    st.session_state.ocr_data = None

uploaded_file = st.sidebar.file_uploader("上传盘点表照片", type=["jpg", "jpeg", "png"])

if uploaded_file is not None:
    st.image(uploaded_file, caption="已上传的纸质表", width=400)
    
    if st.button("🚀 开始识别"):
        with st.spinner("视觉大模型正在看图解析表格..."):
            base64_image = base64.b64encode(uploaded_file.getvalue()).decode('utf-8')
            
            prompt = """
            你是一个工业盘点表识别专家。请仔细识别图片中的表格，提取每一行有效数据，输出一个严格的 JSON 数组。
            数组里的每个对象必须包含三个字段：
            1. "物料名称"：如"卷板"、"矩形管"、"圆钢"、"角钢"、"焊管"等。
            2. "物料描述"：包含所有规格参数，如"T=14"、"180×180×14"、"Φ70"、"60×60×0.79"等。
            3. "预盘点"：该行右侧对应的数字（纯数字）。
            注意：忽略表头，如果某一行没有数据，请不要输出。只输出 JSON 数组，不要输出任何其他解释文字。
            """
            
            try:
                response = client.chat.completions.create(
                    model="glm-4v-flash",
                    messages=[
                        {
                            "role": "user",
                            "content": [
                                {"type": "text", "text": prompt},
                                {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"}}
                            ]
                        }
                    ],
                    temperature=0.1
                )
                result_text = response.choices[0].message.content
                result_text = result_text.replace("```json", "").replace("```", "").strip()
                
                extracted_data = json.loads(result_text)
                
                st.session_state.ocr_data = pd.DataFrame(extracted_data) if extracted_data else pd.DataFrame(columns=["物料名称", "物料描述", "预盘点"])
                
            except Exception as e:
                st.error(f"大模型识别失败，请检查网络或密钥。错误信息：{e}")
                st.session_state.ocr_data = None

if st.session_state.ocr_data is not None:
    st.success("✅ 识别完成！请核对数据（大模型已尽力理解，如有错直接双击修改）：")
    
    edited_df = st.data_editor(st.session_state.ocr_data, use_container_width=True, num_rows="dynamic")
    
    if st.button("📊 智能计算并生成最终报表"):
        final_data = []
        
        code_col = [c for c in db.columns if '编码' in str(c) or '编号' in str(c)]
        code_col = code_col[0] if code_col else ""
        
        for index, row in edited_df.iterrows():
            name = str(row.get("物料名称", ""))
            desc = str(row.get("物料描述", ""))
            qty_str = str(row.get("预盘点", 0))
            
            try:
                qty = float(re.findall(r'\d+\.?\d*', qty_str)[0]) if re.findall(r'\d+\.?\d*', qty_str) else 0.0
            except:
                qty = 0.0
            
            if not name and not desc:
                continue

            dim_nums = [float(n) for n in re.findall(r'\d+\.?\d*', desc)]
            
            actual_val = 0.0
            
            # 核心计算逻辑
            if "板" in name:
                unit = "kg"  # 单位改为 kg
                if len(dim_nums) >= 3:
                    thick, width, length = dim_nums[0], dim_nums[1], dim_nums[2]
                    weight = (thick * width * length / 1000000) * 7.85
                    actual_val = weight * qty
                else:
                    actual_val = qty 
            else:
                unit = "m"   # 单位改为 m
                if len(dim_nums) >= 3:
                    length_mm = dim_nums[-1]
                    single_length = length_mm / 1000.0 if length_mm > 50 else length_mm
                    actual_val = qty * single_length
                else:
                    actual_val = qty 

            # 查库
            matched = pd.DataFrame()
            if not db.empty:
                def check_match(row_data):
                    row_str = str(row_data.values).upper().replace(' ', '')
                    if name and name not in row_str: return False
                    if desc and desc not in row_str: return False
                    return True
                
                mask = db.apply(check_match, axis=1)
                matched = db[mask]

            if not matched.empty:
                raw_code = str(matched.iloc[0][code_col]) if code_col else "未找到"
                m_code = f"'{raw_code}" if raw_code != "nan" else "未匹配到"
            else:
                m_code = "未匹配到"

            # 最终输出表格
            final_data.append({
                "班组": "XXX",
                "盘点位置": f"C{index+1}",
                "物料编码": m_code,
                "物料名称": name,
                "物料描述": desc,
                "预盘点": qty,
                "单重": "",        # 留空
                "重量": "",        # 留空
                "单位": unit,
                "实盘总重": round(actual_val, 3),
                "备注": ""
            })
        
        df_final = pd.DataFrame(final_data)
        
        st.write("### 📊 最终生成的电子盘点表")
        st.dataframe(df_final, use_container_width=True)
        
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False, sheet_name='线边仓盘点表')
        
        st.download_button(
            label="📥 下载标准版 Excel 报表",
            data=buffer.getvalue(),
            file_name="盘点结果报表.xlsx",
            mime="application/vnd.ms-excel"
        )