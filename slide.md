---
marp: true
backgroundColor: #ddd
---

## PythonでExcelファイルを読んだり書いたりするハンズオン

はんなりPython 
2022/06/17

---

## お前誰よ

- 佐野浩士（Hiroshi Sano）[@hrs_sano645](https://twitter.com/hrs_sano645)
  - 🏠:静岡県の富士市🗻
- Job💼
  - [株式会社佐野設計事務所](https://sano-design.info)
    - 自動車向けプレス金型機械の設計事務所
      - 最近は3Dデータの製作/モデリングもやってます！
  - 米農家🌾
- Community🙋
  - 静岡Pythonコミュニティ: Python駿河、PyCon mini Shizuokaスタッフ
  - Code for ふじのくに

---

## FYI

- 静岡Pythonコミュニティの勉強会 Python駿河/Unagi.py
  - 6月は25日開催です！
  - なんとあの有名なパイのメーカーさんの施設で招待制リアル開催になりました🎉
  - 同時生中継します。詳しくはこちら:
    https://py-suruga.connpass.com/event/250842/

---

## 今日のハンズオンでやること

- Pythonで（ライブラリ）パッケージをインストールする
- ExcelなファイルをPythonで読み込む
- Pythonで作った情報をExcelファイルに書き出す

---

## 今日のハンズオンのモチベーション

- 普段手作業でやっていることをPythonでやらせてみる
  - 退屈なことはPy...（by オライリー
- Excelを頑張るのはもうつらい。時間が無い
  - （最近lambda関数やtypescriptを使える楽しさはあるけど）
- 手作業によるミスやヒューマンエラーを防ぐ


<!-- _footer: ヒューマンエラーを無くすのが日々で大事だと思います -->

---

## しくじり先生:マル秘な計算用Excelファイルをメール送付

詳しいことは言わない（言えない）

- Excelで計算したものをPDF化変換してメール送付
- 変換作業は自動化していたが...
メール送付時に不用意にExcelファイルは触らなくていい
  - 一部不備があったExcelファイルをメール送付前に
    **手動で開いて編集** してしまう
- 生成したPDFをメールに添付したつもりが…

<!-- _footer: 後はわかるな… -->

---

まあ1度や2度はやりますよね😇

![irasutoya](https://4.bp.blogspot.com/-L8kmjYNX064/VsGsN2ctx1I/AAAAAAAA39o/NHU8Gnym2GE/s400/kaisya_samui_man.png)

<!-- _footer: 俺みたいになるなよ！ -->

---

## ライブラリを駆使して便利に使う

- Pythonはバッテリーインクルードな言語
- 個人/コミュニティ/企業がサードパーティなライブラリをパッケージ公開する
  - `pip`でインストール, pypiにて公開される
  - パッケージはとっても豊富→エコシステムとして成熟している環境
- Pythonはグルー言語とも言われている
  - ExcelからPDFを生成して、メール添付することもできる
  - Excelにデータを入力してもらってシステムに情報を流し込む

今日はxlsxを読み書きできる `openpyxl` を使います

そのほかのライブラリは適時解説します

---

## 今日使うツール

- Python3.10
  - （仮想環境の）venv上で作業します
- Windows 11
- VS Code

anacondaの方は頑張ってください！

(conda createとかで仮想環境作れるはずであとはあんまり変わらないはずです)

---

## 環境で確認すること

- Python 3.10は入ってる？
  - python3が動くか
  - venvが動くか
- venv上で `pip install openpyxl`
- 今日利用する資料のDL先
  - gitが動く人: 
    `git clone https://github.com/hrsano645/hanpy-handson-202206.git`
  - <https://github.com/hrsano645/hanpy-handson-202206>
    からzipダウンロードして展開

---

## PythonでExcelファイルを読み込む

### やること

- 帳票的なExcelファイルを読み込む: 架空の業務っぽいファイルを使います

---

## PythonでExcelファイルを書き出す

### やること

- 日本の祝日APIからデータを取得する
  - <https://holidays-jp.github.io>
- 集めたデータをExcelファイルへ書き出す

---

## 付録

---

## Excelを扱うほかの方法

- pandas: バックエンドはopenpyxl
- そのほか: こちらに載ってます -> <https://www.python-excel.org>
  - 現在はopenpyxlをおすすめされることが多くなりました
