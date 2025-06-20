#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fam8キャンペーンレポート自動集計システム - CSV読込・統合・データ検証処理
adult/general CSV統合・[total]行除外・エンコーディング自動判定（修正版）
"""

import pandas as pd
from pathlib import Path
import chardet
from loguru import logger
import time


class DataProcessor:
    """CSVデータ処理クラス"""

    def __init__(self, config: dict, target_date_str: str):
        self.config = config
        self.target_date_str = target_date_str
        self.input_dir = Path(config["paths"]["input_dir"]) / target_date_str

        # CSV設定
        self.skip_rows = config["csv_processing"]["skip_header_rows"]
        self.exclude_patterns = config["csv_processing"]["exclude_patterns"]
        self.fallback_encodings = config["csv_processing"]["fallback_encodings"]
        self.chunk_size = config["csv_processing"]["chunk_size"]
        self.large_file_threshold = config["csv_processing"]["large_file_threshold"]

    def process(self) -> pd.DataFrame:
        """CSV統合処理メイン"""
        logger.info("CSV統合処理開始")

        # adult CSV処理
        adult_data = self._process_single_csv("adult")
        logger.info(f"adult CSV処理完了: {len(adult_data)}行")

        # general CSV処理
        general_data = self._process_single_csv("general")
        logger.info(f"general CSV処理完了: {len(general_data)}行")

        # データ統合（adult → general順）
        combined_data = self._combine_data(adult_data, general_data)
        logger.info(f"CSV統合完了: {len(combined_data)}行")

        return combined_data

    def _process_single_csv(self, csv_type: str) -> pd.DataFrame:
        """単一CSV処理"""
        start_time = time.time()

        # ファイルパス取得
        csv_file = self._get_csv_file_path(csv_type)

        # ファイルサイズチェック
        file_size = csv_file.stat().st_size
        logger.info(f"{csv_type} CSVファイルサイズ: {file_size:,} bytes")

        # エンコーディング自動判定（Shift_JIS優先）
        encoding = self._detect_encoding(csv_file)
        logger.info(f"{csv_type} CSV エンコーディング: {encoding}")

        # CSV読込（3行目をヘッダーとして読み込み、列順序・列名は変更しない）
        if file_size > self.large_file_threshold:
            # 大容量ファイル: チャンク読み込み
            data = self._read_large_csv(csv_file, encoding)
        else:
            # 通常ファイル: 一括読み込み
            data = self._read_normal_csv(csv_file, encoding)

        # データクリーニング（[total]行除去のみ）
        cleaned_data = self._clean_data(data, csv_type)

        processing_time = time.time() - start_time
        logger.info(f"{csv_type} CSV処理時間: {processing_time:.2f}秒")

        return cleaned_data

    def _read_normal_csv(self, file_path: Path, encoding: str) -> pd.DataFrame:
        """通常CSV読み込み（3行目をヘッダーとして読み込み）"""
        try:
            # 3行目をヘッダーとして読み込み（skiprows=2で1-2行目をスキップ）
            data = pd.read_csv(
                file_path,
                encoding=encoding,
                skiprows=2,              # 1-2行目をスキップ、3行目がヘッダー
                dtype=str,               # 全て文字列として読み込み
                keep_default_na=False,   # NA値変換無効
                na_filter=False,         # NA値フィルタ無効
                low_memory=False         # メモリ効率より安全性重視
            )

            logger.info(f"CSV読み込み完了: {data.shape[0]}行 × {data.shape[1]}列")
            logger.info(f"実際の列名確認: {list(data.columns)}")

            # 列位置の詳細ログ出力
            self._log_actual_column_positions(data)

            # データサンプル確認
            if not data.empty:
                logger.info(f"データサンプル（最初の3行）:")
                for i in range(min(3, len(data))):
                    sample_row = data.iloc[i].tolist()[:5]  # 最初の5列のみ
                    logger.info(f"  行{i+1}: {sample_row}")

            return data

        except Exception as e:
            logger.error(f"CSV読み込みエラー: {e}")
            raise

    def _log_actual_column_positions(self, data: pd.DataFrame):
        """実際の列位置をログ出力"""
        logger.info("=== 実際のCSV列構造確認 ===")
        
        # 重要な列の実際の位置を特定
        important_columns = ['キャンペーン名', 'Imp', 'Click', 'CV', 'グロス', 'ネット']
        
        for i, col_name in enumerate(data.columns):
            excel_col = self._column_number_to_letter(i + 1)
            if col_name in important_columns:
                logger.info(f"  ★ {col_name} → {excel_col}列（{i+1}番目）")
            else:
                logger.info(f"    {col_name} → {excel_col}列（{i+1}番目）")

    def _clean_data(self, data: pd.DataFrame, csv_type: str) -> pd.DataFrame:
        """データクリーニング（[total]行除去のみ）"""
        original_rows = len(data)

        # 空行除去
        data = data.dropna(how='all')

        # [total]行除去（複数の列をチェック）
        excluded_total = 0
        
        # キャンペーングループ列をチェック
        if 'キャンペーングループ' in data.columns:
            for pattern in self.exclude_patterns:
                mask = data['キャンペーングループ'].astype(str).str.contains(pattern, na=False, case=False, regex=False)

                excluded_rows = mask.sum()
                if excluded_rows > 0:
                    data = data[~mask]
                    excluded_total += excluded_rows
                    logger.info(f"{csv_type} CSV: キャンペーングループ列で'{pattern}'を含む行を{excluded_rows}行除外")

        # キャンペーン名列もチェック
        if 'キャンペーン名' in data.columns:
            for pattern in self.exclude_patterns:
                mask = data['キャンペーン名'].astype(str).str.contains(pattern, na=False, case=False, regex=False)

                excluded_rows = mask.sum()
                if excluded_rows > 0:
                    data = data[~mask]
                    excluded_total += excluded_rows
                    logger.info(f"{csv_type} CSV: キャンペーン名列で'{pattern}'を含む行を{excluded_rows}行除外")

        # インデックスリセット
        data = data.reset_index(drop=True)

        cleaned_rows = len(data)
        logger.info(f"{csv_type} CSV クリーニング: {original_rows}行 → {cleaned_rows}行（{excluded_total}行除外）")

        return data

    def _combine_data(self, adult_data: pd.DataFrame, general_data: pd.DataFrame) -> pd.DataFrame:
        """データ統合（adult → general順、列順序・列名は変更しない）"""

        logger.info(f"統合前データ確認:")
        logger.info(f"  adult: {adult_data.shape[0]}行 × {adult_data.shape[1]}列")
        logger.info(f"  general: {general_data.shape[0]}行 × {general_data.shape[1]}列")

        # 列名統一確認
        adult_columns = list(adult_data.columns)
        general_columns = list(general_data.columns)
        
        if adult_columns != general_columns:
            logger.warning("adult と general で列構成が異なります")
            logger.info(f"adult 列: {adult_columns}")
            logger.info(f"general 列: {general_columns}")
            
            # adultの列順序を基準とする
            base_columns = adult_columns
            
            # general側で不足している列を空文字で補完
            for col in base_columns:
                if col not in general_data.columns:
                    general_data[col] = ""
                    logger.warning(f"general側に不足列'{col}'を空文字で補完")
            
            # adult側で不足している列を空文字で補完
            for col in general_columns:
                if col not in adult_data.columns:
                    adult_data[col] = ""
                    base_columns.append(col)
                    logger.warning(f"adult側に不足列'{col}'を空文字で補完")
            
            # 列順序を統一（adultの順序に合わせる）
            adult_data = adult_data[base_columns]
            general_data = general_data[base_columns]

        logger.info(f"列統一後:")
        logger.info(f"  adult: {adult_data.shape[0]}行 × {adult_data.shape[1]}列")
        logger.info(f"  general: {general_data.shape[0]}行 × {general_data.shape[1]}列")

        # データ統合（adult → general順）
        combined_data = pd.concat([adult_data, general_data], ignore_index=True)

        # 最終データ検証
        self._validate_combined_data(combined_data)

        return combined_data

    def _validate_combined_data(self, data: pd.DataFrame):
        """統合データ検証"""

        # 基本統計
        logger.info("統合データ統計:")
        logger.info(f"  総行数: {len(data):,}行")
        logger.info(f"  総列数: {len(data.columns)}列")
        logger.info(f"  最終列構成: {list(data.columns)}")

        # データサンプル確認
        if not data.empty:
            logger.info("統合データサンプル（最初の3行、重要列のみ）:")
            important_cols = ['キャンペーン名', 'Imp', 'Click', 'CV', 'グロス', 'ネット']
            available_cols = [col for col in important_cols if col in data.columns]
            
            for i in range(min(3, len(data))):
                sample_data = {}
                for col in available_cols:
                    sample_data[col] = data.iloc[i][col]
                logger.info(f"  統合行{i+1}: {sample_data}")

        # 空値チェック
        total_cells = len(data) * len(data.columns)
        empty_cells = data.isnull().sum().sum()
        logger.info(f"  空値セル: {empty_cells:,}個 ({empty_cells/total_cells*100:.2f}%)")

        # 重要列の値チェック
        for col in ['Imp', 'Click', 'CV', 'グロス', 'ネット']:
            if col in data.columns:
                non_empty = data[col].astype(str).str.strip().ne('').sum()
                logger.info(f"  {col}列の非空値: {non_empty}行")

        logger.info("統合データ検証完了")

    def _get_csv_file_path(self, csv_type: str) -> Path:
        """CSVファイルパス取得"""
        if csv_type == "adult":
            filename_template = self.config["files"]["adult_csv"]
        elif csv_type == "general":
            filename_template = self.config["files"]["general_csv"]
        else:
            raise ValueError(f"不正なCSVタイプ: {csv_type}")

        # 日付フォーマット変換 (YYYYMMDD → YYYY-MM-DD)
        date_formatted = f"{self.target_date_str[:4]}-{self.target_date_str[4:6]}-{self.target_date_str[6:]}"
        filename = filename_template.format(date=date_formatted)
        csv_file = self.input_dir / filename

        if not csv_file.exists():
            raise FileNotFoundError(f"{csv_type} CSVファイルが見つかりません: {csv_file}")

        return csv_file

    def _detect_encoding(self, file_path: Path) -> str:
        """エンコーディング自動判定（Shift_JIS優先）"""
        
        # まずShift_JISを試行（日本語CSVの標準）
        try:
            with open(file_path, 'r', encoding='shift_jis') as f:
                f.read(1024)  # 1KB試し読み
            logger.info("エンコーディング: shift_jis (優先試行成功)")
            return 'shift_jis'
        except UnicodeDecodeError:
            logger.debug("shift_jis読み込み失敗、自動判定に移行")
        
        # 自動判定
        try:
            with open(file_path, 'rb') as f:
                raw_data = f.read(10240)  # 10KB読み取り

            result = chardet.detect(raw_data)
            detected_encoding = result['encoding']
            confidence = result['confidence']

            logger.debug(f"エンコーディング判定結果: {detected_encoding} (信頼度: {confidence:.2f})")

            # 信頼度が高い場合は採用
            if confidence >= 0.7:
                return detected_encoding
            else:
                logger.warning(f"エンコーディング判定信頼度が低い: {confidence:.2f}")
                return self._try_fallback_encodings(file_path)

        except Exception as e:
            logger.warning(f"エンコーディング自動判定失敗: {e}")
            return self._try_fallback_encodings(file_path)

    def _try_fallback_encodings(self, file_path: Path) -> str:
        """フォールバックエンコーディング試行"""
        for encoding in self.fallback_encodings:
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    f.read(1024)  # 1KB試し読み
                logger.info(f"フォールバックエンコーディング成功: {encoding}")
                return encoding
            except UnicodeDecodeError:
                continue

        # 全て失敗した場合はutf-8で強制読み込み
        logger.warning("全てのエンコーディング試行が失敗、utf-8で強制読み込み")
        return "utf-8"

    def _read_large_csv(self, file_path: Path, encoding: str) -> pd.DataFrame:
        """大容量CSV読み込み（チャンク処理）"""
        logger.info("大容量ファイル検出、チャンク読み込み開始")

        chunks = []
        try:
            chunk_reader = pd.read_csv(
                file_path,
                encoding=encoding,
                skiprows=2,              # 3行目をヘッダーとして使用
                dtype=str,
                keep_default_na=False,
                na_filter=False,
                chunksize=self.chunk_size
            )

            for i, chunk in enumerate(chunk_reader):
                chunks.append(chunk)
                if i % 10 == 0:  # 10チャンクごとにログ出力
                    logger.debug(f"チャンク処理中: {i+1}チャンク完了")

            data = pd.concat(chunks, ignore_index=True)
            logger.info(f"チャンク読み込み完了: {len(chunks)}チャンク")
            return data

        except Exception as e:
            logger.error(f"チャンク読み込みエラー: {e}")
            raise

    def _column_number_to_letter(self, col_num: int) -> str:
        """列番号をアルファベットに変換"""
        result = ""
        while col_num > 0:
            col_num -= 1
            result = chr(65 + (col_num % 26)) + result
            col_num //= 26
        return result