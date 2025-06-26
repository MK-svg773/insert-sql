import pandas as pd
from datetime import datetime

# Excelファイルを読み込み、全シートの列名から前後の空白を削除して格納
excel_path = "販売管理データ.xlsx"
xls = pd.ExcelFile(excel_path)
data = {name: xls.parse(name).rename(columns=lambda x: str(x).strip()) for name in xls.sheet_names}

# 組織データの抽出と重複排除
org_df = data["組織・社員"][["組織コード", "本部名称", "部名称", "課名称"]].drop_duplicates()
org_inserts = []
for _, row in org_df.iterrows():
    values = [
        f"'{row['組織コード']}'",
        f"'{row['本部名称']}'" if pd.notna(row['本部名称']) else "NULL",
        f"'{row['部名称']}'" if pd.notna(row['部名称']) else "NULL",
        f"'{row['課名称']}'" if pd.notna(row['課名称']) else "NULL",
    ]
    sql = f"INSERT INTO 組織 (組織コード, 本部名称, 部名称, 課名称) VALUES ({', '.join(values)});"
    org_inserts.append(sql)

# 社員データの処理
emp_df = data["組織・社員"][["社員番号", "氏名", "職種", "職位", "EMAIL", "組織コード"]].dropna(subset=["社員番号"])
emp_df["社員番号"] = emp_df["社員番号"].astype(int).astype(str).str.zfill(6)
emp_inserts = []
for _, row in emp_df.iterrows():
    values = [
        f"'{row['社員番号']}'",
        f"'{row['氏名']}'",
        f"'{row['職種']}'" if pd.notna(row['職種']) else "NULL",
        f"'{row['職位']}'" if pd.notna(row['職位']) else "NULL",
        f"'{row['EMAIL']}'" if pd.notna(row['EMAIL']) else "NULL",
        f"'{row['組織コード']}'",
    ]
    sql = f"INSERT INTO 社員 (社員番号, 氏名, 職種, 職位, メールアドレス, 組織コード) VALUES ({', '.join(values)});"
    emp_inserts.append(sql)

# 顧客データ
customer_df = data["顧客"]
customer_df["顧客ID"] = customer_df["顧客ID"].astype(str).str.zfill(7)
customer_inserts = []
for _, row in customer_df.iterrows():
    values = [
        f"'{row['顧客ID']}'",
        f"'{row['顧客名称']}'",
        f"'{row['住所']}'",
        f"'{str(row['電話番号']).split('.')[0]}'",
        f"'{row['担当課']}'"
    ]
    sql = f"INSERT INTO 顧客 (顧客ID, 顧客名称, 住所, 電話番号, 担当課) VALUES ({', '.join(values)});"
    customer_inserts.append(sql)

# 商品タイプコード
product_sheet = data["商品"]
type_df = product_sheet[["商品タイプコード"]].drop_duplicates()
type_inserts = [
    f"INSERT INTO 商品タイプ (商品タイプコード) VALUES ('{row['商品タイプコード']}');"
    for _, row in type_df.iterrows() if pd.notna(row['商品タイプコード'])
]

# スポーツ種別
sport_df = product_sheet[["スポーツ種別コード", "スポーツ名称"]].drop_duplicates()
sport_inserts = [
    f"INSERT INTO スポーツ種別 (スポーツ種別コード, スポーツ名称) VALUES ('{row['スポーツ種別コード']}', '{row['スポーツ名称']}');"
    for _, row in sport_df.iterrows() if pd.notna(row['スポーツ種別コード'])
]

# 商品データ
product_df = product_sheet[["商品ID", "商品名称", "商品タイプコード", "スポーツ種別コード", "標準単価", "販売可否"]].copy()
product_df["商品ID"] = product_df["商品ID"].astype(str).str.zfill(10)
product_inserts = []
for _, row in product_df.iterrows():
    price = row['標準単価']
    sale_flag = row['販売可否']
    values = [
        f"'{row['商品ID']}'",
        f"'{row['商品名称']}'",
        f"'{row['商品タイプコード']}'",
        f"'{row['スポーツ種別コード']}'",
        str(int(price)) if pd.notna(price) and str(price).strip() != "" else "NULL",
        str(int(sale_flag)) if pd.notna(sale_flag) and str(sale_flag).strip() != "" else "NULL"
    ]
    sql = f"INSERT INTO 商品 (商品ID, 商品名称, 商品タイプコード, スポーツ種別コード, 標準単価, 販売可否) VALUES ({', '.join(values)});"
    product_inserts.append(sql)


# 受注テーブル
order_df = data["受注"]
order_df["受注ID"] = order_df["受注ID"].astype(str)
order_df["社員番号"] = order_df["社員番号"].astype(int).astype(str).str.zfill(6)
order_df["顧客ID"] = order_df["顧客ID"].astype(str).str.zfill(7)
order_inserts = []
for _, row in order_df.iterrows():
    values = [
        f"'{row['受注ID']}'",
        f"'{row['受注年月日'].strftime('%Y-%m-%d')}'" if pd.notna(row['受注年月日']) else "NULL",
        f"'{row['顧客ID']}'",
        f"'{row['社員番号']}'",
        f"'{row['発送年月日'].strftime('%Y-%m-%d')}'" if pd.notna(row['発送年月日']) else "NULL",
    ]
    sql = f"INSERT INTO 受注 (受注ID, 受注年月日, 顧客ID, 社員番号, 発送年月日) VALUES ({', '.join(values)});"
    order_inserts.append(sql)


# 受注明細（受注シートから取得）
# 明細情報が '受注' シート内に含まれている前提で処理
if "受注" in data:
    odetail_df = data["受注"]
    required_columns = ["受注ID", "商品ID", "受注数量", "販売単価"]
    missing_cols = [col for col in required_columns if col not in odetail_df.columns]
    if missing_cols:
        print(f"エラー: 受注明細の生成に必要な列が不足しています: {missing_cols}")
        odetail_inserts = []
    else:
        odetail_df = odetail_df[required_columns].copy()
        odetail_df["受注ID"] = odetail_df["受注ID"].astype(str)
        odetail_df["商品ID"] = odetail_df["商品ID"].astype(str).str.zfill(11)
        odetail_inserts = []
        for _, row in odetail_df.iterrows():
            values = [
                f"'{row['受注ID']}'",
                f"'{row['商品ID']}'",
                str(int(row['受注数量'])),
                str(int(row['販売単価']))
            ]
            sql = f"INSERT INTO 受注明細 (受注ID, 商品ID, 受注数量, 販売単価) VALUES ({', '.join(values)});"
            odetail_inserts.append(sql)
else:
    odetail_inserts = []
    print("エラー: シート '受注' が見つかりません。")



# 社員目標
emp_target_df = data["社員目標"]
emp_target_df["社員番号"] = emp_target_df["社員番号"].astype(int).astype(str).str.zfill(6)
emp_target_inserts = []
for _, row in emp_target_df.iterrows():
    values = [
        f"'{row['社員番号']}'",
        f"'{row['組織コード']}'",
    ] + [str(int(row[col])) for col in emp_target_df.columns if '目標' in col]
    sql = f"INSERT INTO 社員目標 VALUES ({', '.join(values)});"
    emp_target_inserts.append(sql)

# 出力ファイル保存
all_inserts = (
    org_inserts + emp_inserts + customer_inserts +
    type_inserts + sport_inserts + product_inserts +
    order_inserts + odetail_inserts + emp_target_inserts
)
with open("販売管理データ一括登録.sql", "w", encoding="utf-8") as f:
    for stmt in all_inserts:
        f.write(stmt + "\n")

