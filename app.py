import streamlit as st
import json
import io
from pptx import Presentation
from pptx.util import Inches

st.set_page_config(page_title="MedConcept PPTX生成", layout="wide")

# --- 補助関数：図形内のテキストをフォント維持で置換 ---
def replace_text_in_shape(shape, replacements):
    if not shape.has_text_frame:
        return
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            for old_text, new_text in replacements.items():
                if old_text in run.text:
                    # フォントスタイルを維持したまま文字だけ置換
                    run.text = run.text.replace(old_text, str(new_text))

# --- 補助関数：不要な図形を完全に削除 ---
def delete_shape(shape):
    sp_tree = shape.element.getparent()
    sp_tree.remove(shape.element)

# --- メイン処理関数 ---
def generate_pptx(json_data, uploaded_images):
    # テンプレートの読み込み
    prs = Presentation("template.pptx")

    # 1. テキストの置換マッピングを作成
    brand_info = f"カラー：{json_data.get('brandColors', '')}\nブランドイメージ：{'、'.join(json_data.get('brandImages', []))}"
    
    replacements = {
        "{{productName}}": json_data.get("productName", ""),
        "{{itemName}}": json_data.get("itemName", ""),
        "{{spec}}": json_data.get("spec", ""),
        "{{target}}": json_data.get("target", ""),
        "{{scene}}": json_data.get("scene", ""),
        "{{objectiveA}}": json_data.get("objectiveA", ""),
        "{{objectiveB}}": json_data.get("objectiveB", ""),
        "{{before}}": json_data.get("before", ""),
        "{{after}}": json_data.get("after", ""),
        "{{concept}}": json_data.get("concept", ""),
        "{{brandInfo}}": brand_info,
        "{{designExterior}}": "、".join(json_data.get("designExterior", [])),
        "{{functional}}": "、".join(json_data.get("functional", [])),
        "{{toneManner}}": "\n".join(json_data.get("toneManner", [])),
    }

    # 変化タイプ（最大4つ）のマッピング
    cb = json_data.get("changeTypesBefore", [])
    ca = json_data.get("changeTypesAfter", [])
    
    for i in range(4):
        replacements[f"{{{{cb{i+1}}}}}"] = cb[i] if i < len(cb) else ""
        replacements[f"{{{{ca{i+1}}}}}"] = ca[i] if i < len(ca) else ""

    # 2. 全スライドのテキスト置換と不要図形の削除
    for slide in prs.slides:
        shapes_to_delete = []
        
        # グループ化された図形も再帰的にチェックする内部関数
        def process_shapes(shapes):
            for shape in shapes:
                if shape.shape_type == 6: # グループ図形
                    process_shapes(shape.shapes)
                elif shape.has_text_frame:
                    # 4つ目がない場合、{{cb4}}や{{ca4}}を含む図形を削除候補に追加
                    delete_flag = False
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            if "{{cb4}}" in run.text and len(cb) < 4:
                                delete_flag = True
                            if "{{ca4}}" in run.text and len(ca) < 4:
                                delete_flag = True
                    
                    if delete_flag:
                        shapes_to_delete.append(shape)
                    else:
                        replace_text_in_shape(shape, replacements)
                        
                elif shape.has_table: # テーブル内のテキスト置換
                    for row in shape.table.rows:
                        for cell in row.cells:
                            replace_text_in_shape(cell, replacements)

        process_shapes(slide.shapes)

        # マークした図形を削除
        for shape in shapes_to_delete:
            try:
                delete_shape(shape)
            except Exception:
                pass

    # 3. 画像のレイアウト配置（スライド6〜10 / インデックス5〜9）
    # A案=5, B案=6, C案=7, D案=8, E案=9
    slide_indices = {"A案": 5, "B案": 6, "C案": 7, "D案": 8, "E案": 9}
    
    # 2行3列のグリッド計算用の設定（16:9スライド基準）
    margin_x, margin_y = Inches(0.5), Inches(1.5)
    cell_w, cell_h = Inches(3.0), Inches(2.0)
    cols = 3

    for plan_name, images in uploaded_images.items():
        if plan_name in slide_indices and len(prs.slides) > slide_indices[plan_name]:
            slide = prs.slides[slide_indices[plan_name]]
            
            for idx, img_file in enumerate(images[:6]): # 最大6枚まで
                row = idx // cols
                col = idx % cols
                x = margin_x + (col * cell_w)
                y = margin_y + (row * cell_h)
                
                img_stream = io.BytesIO(img_file.read())
                try:
                    # widthだけ指定し、アスペクト比を自動維持して挿入
                    slide.shapes.add_picture(img_stream, x, y, width=cell_w - Inches(0.2))
                except Exception as e:
                    st.warning(f"{plan_name}の画像挿入に失敗しました: {e}")

    # 4. メモリ上に保存して出力
    ppt_stream = io.BytesIO()
    prs.save(ppt_stream)
    ppt_stream.seek(0)
    return ppt_stream

# --- UI構築 ---
st.title("MedConcept - 企画書PPTX自動生成")
st.markdown("STEP7の画像とSTEP8のJSONデータを入力して、パワポを生成します。")

# STEP 7: 画像アップロード
st.header("STEP 7: 画像アップロード (各案5〜6枚推奨)")
uploaded_images = {}
cols = st.columns(5)
plans = ["A案", "B案", "C案", "D案", "E案"]

for i, plan in enumerate(plans):
    with cols[i]:
        st.subheader(plan)
        uploaded_images[plan] = st.file_uploader(f"{plan}の画像", accept_multiple_files=True, type=["png", "jpg", "jpeg"], key=plan)

# STEP 8: JSONテキスト入力
st.header("STEP 8: JSONデータ入力")
json_text = st.text_area("HTMLアプリで生成されたJSONデータを貼り付けてください", height=300)

if st.button("📊 企画書を作成", type="primary"):
    if not json_text.strip():
        st.error("JSONデータを入力してください。")
    else:
        try:
            # JSONのパース
            json_data = json.loads(json_text)
            
            with st.spinner("PowerPointを生成中..."):
                ppt_stream = generate_pptx(json_data, uploaded_images)
                
            st.success("PowerPointの生成が完了しました！")
            st.download_button(
                label="📥 proposal.pptx をダウンロード",
                data=ppt_stream,
                file_name=f"proposal_{json_data.get('itemName', 'untitled')}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
            
        except json.JSONDecodeError:
            st.error("JSONのフォーマットが正しくありません。コピー忘れや余分な文字がないか確認してください。")
        except Exception as e:
            st.error(f"エラーが発生しました: {e}")