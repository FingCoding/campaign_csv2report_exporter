#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fam8キャンペーンレポート自動集計システム - メイン制御・工程管理
処理フロー制御・エラーハンドリング・ログ管理統括クラス（修正版）
"""

import sys
import shutil
from pathlib import Path
from datetime import datetime, timedelta
import tomli
from loguru import logger
import psutil
import time

from data_processor import DataProcessor
from data_handler import DataHandler
from format_manager import FormatManager


class CampaignReportOrchestrator:
    """fam8キャンペーンレポート自動集計メイン制御クラス"""

    def __init__(self, debug_mode: bool = False):
        self.debug_mode = debug_mode
        self.config = None
        self.target_date = None
        self.target_date_str = None
        self.start_time = time.time()

    @logger.catch
    def execute(self, target_date: str = None):
        """メイン処理実行"""
        try:
            # 工程1: 設定ファイル読込
            self._load_config()

            # 工程2: 処理対象日計算・設定
            self._calculate_target_date(target_date)

            # 工程3: ログ初期化
            self._initialize_logging()

            logger.info("="*60)
            logger.info(f"fam8キャンペーンレポート自動集計開始")
            logger.info(f"処理対象日: {self.target_date_str}")
            logger.info(f"デバッグモード: {self.debug_mode}")
            logger.info("="*60)

            # 工程4: 環境バリデーション
            self._validate_environment()

            # 工程5: CSV統合・集計処理
            self._process_csv_data()

            # 工程6: Excel出力処理（データ貼付→関数埋込→書式設定の順序保証）
            self._build_excel_report()

            # 工程7: ファイル配布
            self._distribute_files()

            # 工程8: 処理完了ログ
            self._log_completion()

            logger.info("="*60)
            logger.info("fam8キャンペーンレポート自動集計完了")
            logger.info("="*60)

        except Exception as e:
            logger.error(f"致命的エラー発生: {e}")
            sys.exit(1)

    def _load_config(self):
        """設定ファイル読込"""
        config_path = Path("config.toml")
        if not config_path.exists():
            raise FileNotFoundError(f"設定ファイルが見つかりません: {config_path}")

        with open(config_path, "rb") as f:
            self.config = tomli.load(f)

        logger.info(f"設定ファイル読込完了: {config_path}")

    def _calculate_target_date(self, target_date: str = None):
        """処理対象日計算"""
        if target_date:
            try:
                self.target_date = datetime.strptime(target_date, "%Y%m%d")
                self.target_date_str = target_date
            except ValueError:
                raise ValueError(f"日付形式が正しくありません: {target_date} (YYYYMMDD形式で入力)")
        else:
            # 前日を自動計算
            self.target_date = datetime.now() - timedelta(days=1)
            self.target_date_str = self.target_date.strftime("%Y%m%d")

        logger.info(f"処理対象日設定完了: {self.target_date_str}")

    def _initialize_logging(self):
        """ログ初期化"""
        # ログディレクトリ作成
        log_dir = Path("log") / self.target_date_str
        log_dir.mkdir(parents=True, exist_ok=True)

        # ログファイルパス
        log_file = log_dir / f"{self.target_date_str}.log"
        performance_log = log_dir / f"{self.target_date_str}_performance.log"

        # ログレベル設定
        log_level = "DEBUG" if self.debug_mode else "INFO"

        # ログ設定
        logger.remove()  # デフォルトハンドラー削除

        # コンソールログ
        logger.add(
            sys.stdout,
            level=log_level,
            format="<green>{time:YYYY-MM-DD HH:mm:ss}</green> | <level>{level: <8}</level> | <level>{message}</level>"
        )

        # ファイルログ（追記方式）
        logger.add(
            str(log_file),
            level="DEBUG",
            format="[{level}] {time:YYYY-MM-DD HH:mm:ss} → {message}",
            mode="a",
            rotation="10 MB",
            retention="30 days"
        )

        logger.info(f"ログ初期化完了: {log_file}")

    def _validate_environment(self):
        """環境バリデーション（修正版）"""
        logger.info("環境バリデーション開始")

        # 入力CSVファイル存在チェック
        input_dir = Path(self.config["paths"]["input_dir"]) / self.target_date_str

        # 正確な日付フォーマット使用（YYYY-MM-DD）
        date_formatted = f"{self.target_date_str[:4]}-{self.target_date_str[4:6]}-{self.target_date_str[6:]}"
        
        adult_csv = input_dir / self.config["files"]["adult_csv"].format(date=date_formatted)
        general_csv = input_dir / self.config["files"]["general_csv"].format(date=date_formatted)

        logger.info(f"CSVファイル存在確認:")
        logger.info(f"  adult CSV: {adult_csv}")
        logger.info(f"  general CSV: {general_csv}")

        if not adult_csv.exists():
            raise FileNotFoundError(f"adult CSVファイルが見つかりません: {adult_csv}")
        if not general_csv.exists():
            raise FileNotFoundError(f"general CSVファイルが見つかりません: {general_csv}")

        # ファイルサイズ確認
        adult_size = adult_csv.stat().st_size
        general_size = general_csv.stat().st_size
        logger.info(f"CSVファイルサイズ:")
        logger.info(f"  adult CSV: {adult_size:,} bytes")
        logger.info(f"  general CSV: {general_size:,} bytes")

        # FilterInput_Csvreport.xlsx存在チェック
        filter_excel = Path(self.config["paths"]["filter_input_excel"])
        if not filter_excel.exists():
            raise FileNotFoundError(f"FilterInput_Csvreport.xlsxが見つかりません: {filter_excel}")

        logger.info(f"FilterInput_Csvreport.xlsx確認完了: {filter_excel}")

        # 出力ディレクトリ作成
        output_dir = Path(self.config["paths"]["output_dir"]) / self.target_date_str
        output_dir.mkdir(parents=True, exist_ok=True)
        logger.info(f"出力ディレクトリ作成完了: {output_dir}")

        logger.info("環境バリデーション完了")

    def _process_csv_data(self):
        """CSV統合・集計処理"""
        logger.info("CSV統合・集計処理開始")

        processor = DataProcessor(self.config, self.target_date_str)
        combined_data = processor.process()

        self.combined_csv_data = combined_data
        logger.info(f"CSV統合完了: {len(combined_data)}行")

        # 統合データの基本情報をログ出力
        logger.info(f"統合データ詳細:")
        logger.info(f"  行数: {len(combined_data):,}行")
        logger.info(f"  列数: {len(combined_data.columns)}列")
        logger.info(f"  列構成: {list(combined_data.columns)}")

    def _build_excel_report(self):
        """Excel出力処理（順序保証：データ貼付→関数埋込→書式設定）"""
        logger.info("Excel出力処理開始")

        # データ操作（CSV貼付＋関数埋込）
        data_handler = DataHandler(self.config, self.target_date_str)
        workbook = data_handler.process(self.combined_csv_data)

        # 書式設定（関数埋込後に実行）
        format_manager = FormatManager(self.config)
        format_manager.apply_formatting(workbook)

        # ファイル保存
        data_handler.save_workbook(workbook)

        logger.info("Excel出力処理完了")

    def _distribute_files(self):
        """ファイル配布"""
        logger.info("ファイル配布開始")

        # 元ファイル
        source_file = Path(self.config["paths"]["filter_input_excel"])

        # 配布先ディレクトリ
        output_dir = Path(self.config["paths"]["output_dir"]) / self.target_date_str

        # 配布先ファイル名（YYYYMMDD形式）
        output_filename = self.config["files"]["output_filename"].format(date=self.target_date_str)
        output_file = output_dir / output_filename

        # ファイルコピー
        shutil.copy2(source_file, output_file)

        logger.info(f"ファイル配布完了:")
        logger.info(f"  元ファイル: {source_file}")
        logger.info(f"  配布先: {output_file}")

    def _log_completion(self):
        """処理完了ログ"""
        end_time = time.time()
        processing_time = end_time - self.start_time

        # メモリ使用量
        memory_usage = psutil.Process().memory_info().rss / 1024 / 1024  # MB

        logger.info("="*40)
        logger.info("処理完了統計")
        logger.info(f"処理時間: {processing_time:.2f}秒")
        logger.info(f"メモリ使用量: {memory_usage:.2f}MB")
        logger.info(f"処理対象日: {self.target_date_str}")
        logger.info(f"CSV統合行数: {len(self.combined_csv_data):,}行")
        logger.info("="*40)