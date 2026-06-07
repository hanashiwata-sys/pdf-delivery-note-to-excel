import pdfplumber
import re
import glob
import pandas as pd
import os

pdf_files = glob.glob("PDF(納品書)/*.pdf")

all_data = []

for pdf_path in pdf_files:
    
    print(f"処理中PDF: {pdf_path}")
    
    # PDF名取得
    pdf_name = os.path.basename(pdf_path)
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            
            current_date = ""
            company_name = ""
            
            for page in pdf.pages:
                
                text = page.extract_text()
                
                if not text:
                    continue
                
                lines = text.split("\n")
                
                # 会社名取得（2行目）
                if len(lines) > 0:
                    company_name = lines[1]
                
                for line in lines:
                    
                    # -----------------------------
                    # 日付抽出
                    # -----------------------------
                    date_match = re.search(
                        r"(\d+月\s+\d+\s+日)",
                        line
                    )
                    
                    if date_match:
                        current_date = date_match.group(1)
                    
                    # -----------------------------
                    # 行分割
                    # -----------------------------
                    parts = line.split()
                    
                    # 要素数不足は除外
                    if len(parts) < 5:
                        continue
                    
                    # -----------------------------
                    # 後ろ2つが数字か確認
                    # 金額・単価
                    # -----------------------------
                    if (
                        parts[-1].isdigit()
                        and parts[-2].isdigit()
                        and parts[-4].isdigit()
                    ):
                        
                        try:
                            amount = int(parts[-1])
                            price = int(parts[-2])
                            quantity = int(parts[-4])
                            
                            # 単位
                            unit = parts[-3]
                            
                            # 商品名
                            product_name = " ".join(parts[:-4])
                            
                            # 不要行除外
                            exclude_words = [
                                "商品名",
                                "合計金額",
                                "納品書",
                                "備考"
                            ]
                            
                            if any(
                                word in product_name
                                for word in exclude_words
                            ):
                                continue
                            
                            data = {
                                "PDF名": pdf_name,
                                "会社名": company_name,
                                "日付": current_date,
                                "商品名": product_name,
                                "数量": quantity,
                                "単位": unit,
                                "単価": price,
                                "金額": amount
                            }
                            
                            all_data.append(data)
                        
                        except:
                            print(f"変換エラー: {line}")
    
    except Exception as e:
        print(f"PDF読み込みエラー: {pdf_path}")
        print(e)

# ---------------------------------
# DataFrame化
# ---------------------------------
df = pd.DataFrame(all_data)

# ---------------------------------
# Excel出力
# ---------------------------------
df.to_excel("納品書一覧.xlsx", index=False)

# ---------------------------------
# CSV出力
# ---------------------------------
df.to_csv(
    "納品書一覧.csv",
    index=False,
    encoding="utf-8-sig"
)

# ---------------------------------
# 商品別集計
# ---------------------------------
summary_df = (
    df.groupby("商品名")["金額"]
    .sum()
    .sort_values(ascending=False)
    .reset_index()
)

summary_df.to_excel(
    "商品別集計.xlsx",
    index=False
)

print("処理完了")