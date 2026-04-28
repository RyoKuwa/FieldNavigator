// 調査用地図 v1 設定ファイル
// ------------------------------------------------------------
// スプレッドシートはアプリ画面の「追加」から端末ごとに登録します。
// 登録情報はブラウザの localStorage に保存されます。

window.FIELD_MAP_CONFIG = {
  TITLE: "調査用地図",

  // 新規登録時の既定シート名。日本語環境のGoogleスプレッドシートでは通常「シート1」です。
  DEFAULT_SHEET_NAME: "シート1",

  // 端末内保存に使うキー。通常は変更不要です。
  LOCAL_DATASETS_STORAGE_KEY: "fieldMap.localDatasets.v1",
  ACTIVE_DATASET_STORAGE_KEY: "fieldMap.activeDatasetId.v1",

  // 初期表示。AUTO_FIT_BOUNDS が true の場合、読み込み後に全地点へ自動ズームします。
  START_CENTER: [133.0, 33.7], // [longitude, latitude]
  START_ZOOM: 7,
  AUTO_FIT_BOUNDS: true,

  // 背景地図の初期値。"detail" は航空写真、"pale" は地理院淡色地図です。
  INITIAL_BASEMAP: "detail",

  // 画面上でこのピクセル数以内にある記録を近接記録としてまとめます。
  PROXIMITY_PIXELS: 10,

  // 同定色の割り当て時に、各記録から最も近い別同定を何種類まで見るか。
  // 固定距離ではなく「近い順」のため、ズームや地図表示範囲には依存しません。
  COLOR_NEAREST_DIFFERENT_TAXA: 5,

  // 表示に使う列名。
  COLUMNS: {
    id: "記録ID",
    taxon: "同定",
    recordType: "記録の種類",
    repository: "所蔵",
    specimenId: "標本ID",
    locality: "場所",
    latitude: "Latitude",
    longitude: "Longitude",
    year: "年",
    month: "月",
    day: "日",
    date: "表示する日付",
    remarks: "備考"
  }
};
