// 調査結果地図 v1 設定ファイル
// ------------------------------------------------------------
// スプレッドシートはアプリ画面の「追加」から端末ごとに登録します。
// 登録情報はブラウザの localStorage に保存されます。

window.FIELD_MAP_CONFIG = {
  TITLE: "調査結果地図",

  // 新規登録時の既定シート名。日本語環境のGoogleスプレッドシートでは通常「シート1」です。
  DEFAULT_SHEET_NAME: "シート1",

  // 端末内保存に使うキー。通常は変更不要です。
  LOCAL_DATASETS_STORAGE_KEY: "fieldMap.localDatasets.v1",
  ACTIVE_DATASET_STORAGE_KEY: "fieldMap.activeDatasetId.v1",
  FILTERS_STORAGE_KEY_PREFIX: "fieldMap.filters.v1",

  // 初期表示。AUTO_FIT_BOUNDS が true の場合、読み込み後に全地点へ自動ズームします。
  START_CENTER: [133.0, 33.7], // [longitude, latitude]
  START_ZOOM: 7,
  AUTO_FIT_BOUNDS: true,
  FOCUS_SINGLE_MARKER_ZOOM: 14,
  FOCUS_SINGLE_MARKER_MAX_ZOOM: 16,

  // 背景地図の初期値。"detail" は航空写真、"pale" は地理院淡色地図です。
  INITIAL_BASEMAP: "detail",

  // 画面上でこのピクセル数以内にある記録を近接記録としてまとめます。
  PROXIMITY_PIXELS: 10,

  // ポップアップの前後ボタンで切り替える近接マーカーの判定半径です。
  // detail.zip と同じく、近接記録の統合距離と同程度の 10 px を初期値にします。
  // 地理的距離ではなく、現在の画面上のピクセル距離で判定します。
  POPUP_NEARBY_MARKER_PIXELS: 10,

  // 読み込みに時間がかかる場合の案内とタイムアウトです。
  LOAD_SLOW_NOTICE_MS: 10000,
  LOAD_TIMEOUT_MS: 30000,

  // 色分け時に、各記録から最も近い別の値を何種類まで見るか。
  // 固定距離ではなく「近い順」のため、ズームや地図表示範囲には依存しません。
  COLOR_NEAREST_DIFFERENT_TAXA: 5,

  // 表示に使う列名。
  COLUMNS: {
    id: "記録ID",
    taxon: "分類群",
    recordType: "記録の種類",
    repository: "所蔵",
    specimenId: "標本ID",
    locality: "場所",
    latitude: "緯度",
    longitude: "経度",
    year: "年",
    month: "月",
    day: "日",
    date: "日付",
    remarks: "備考"
  }
};
