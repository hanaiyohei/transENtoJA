# Livelike 楽天 価格見直し試算

入力（Google Drive からxlsxで書き出し、同じフォルダに置く。リポジトリには含めない）:
- `tsugou.xlsx` : 総勘定元帳_AE_販売_突合表（AEordersシート）
- `daily.xlsx`  : dailydata（salesシート）

手順:
1. `python3 01_bundles.py` : AE仕入と楽天受注を受注単位に集計（bundles.pkl）
2. 単品注文を抽出して single.pkl を作成（01の出力から n_ord==1 & n_ae==1）
3. `python3 02_reprice_sim.py` : 価格式 `(AE仕入+300)/(1-0.18)*2.5` による仕入帯別の推奨価格と弾力性シナリオ

前提: 入金ベースの控除率24%（比例分18% + 固定分約6%の推定）、1件あたり固定費¥300。
