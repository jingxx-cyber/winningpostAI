# Implementation Plan: AI Secretary v2 Refactoring

システムの軽量化・高速化に向けた大規模リファクタリングの計画です。

## Proposed Changes

### [Logic Components]
- **[NEW] logic/state_manager.py**: 画面から得られた数値や状況を保持する `State` クラス。
- **[NEW] logic/event_detector.py**: Stateの変化を監視し、`RACE_RESULT` などのイベントを発火させる。
- **[MODIFY] logic/capture_thread.py**: 
    - メインループからGemini呼び出しを削除。
    - シーン判定後に、特定の座標のみを `pytesseract` に渡すよう変更。
    - ユーザー入力があった時のみ [GeminiBrowser](file:///c:/Users/zebel/.gemini/antigravity/scratch/winpost_capture/capture.py#112-196) を通じて発話。

### [Rule Sets]
- **[NEW] config/ocr_regions.json**: Winning Post 10 2025 の各画面における、OCR対象領域の座標データ。

## Verification Plan

### Automated Tests
- 各シーンのキャプチャ画像に対して、限定領域OCRが正しく文字を取得できるか確認。
- Stateの変更が正しく検知され、イベントがトリガーされるか確認。

### Manual Verification
- 常時ループ中にAPIリクエストが発生していないことをネットワークログ/デバッグログで確認。
- ユーザー質問時に、Stateの情報が正しくGeminiに伝わっているかプロンプトを確認。
