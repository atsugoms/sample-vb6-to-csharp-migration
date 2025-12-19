# クイックスタートガイド (Quick Start Guide)

## 必要な環境

- Windows OS (Windows 7以降)
- Visual Basic 6.0

## ビルド手順

1. Visual Basic 6.0を起動
2. ファイル → プロジェクトを開く
3. `WarehouseSystem.vbp` を選択
4. F5キーで実行（デバッグモード）

または

4. ファイル → WarehouseSystem.exeの作成
5. 保存先を指定してコンパイル

## プロジェクト構成

```
WarehouseSystem/
├── WarehouseSystem.vbp          - プロジェクトファイル
├── modDatabase.bas              - データベース管理
├── cls*.cls                     - データモデル（3ファイル）
└── frm*.frm                     - 画面フォーム（4ファイル）
```

## 実装済み機能

### メインメニュー (frmMain)
- 入庫予定管理へのリンク
- 入庫実績管理へのリンク
- ロケーション管理へのリンク

### 入庫予定管理 (frmScheduledReceiving)
- 一覧表示
- 追加・編集・削除

### 入庫実績管理 (frmActualReceiving)
- 一覧表示
- 追加・編集・削除
- バーコード種類、検品エラー、棚番号の記録

### ロケーション管理 (frmLocation)
- 一覧表示
- 追加・編集・削除
- 棚属性、割り当てルールの設定

## データ構造

### 入庫予定 (clsScheduledReceiving)
- ID, 発注番号, 商品コード, 数量, 納期

### 入庫実績 (clsActualReceiving)
- ID, 発注番号, 商品コード, 数量, バーコード種類, 検品エラー, 棚番号, 入庫日

### ロケーション (clsLocation)
- ID, 棚番号, 棚属性, 割り当てルール

## 注意事項

- データはメモリ上のみに保存（アプリ終了で消失）
- サンプルデータは起動時に自動作成
- セキュリティ機能は未実装（サンプルのため）

詳細は README.md を参照してください。
