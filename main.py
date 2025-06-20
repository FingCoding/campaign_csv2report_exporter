#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fam8キャンペーンレポート自動集計システム - メインエントリーポイント
処理対象: 前日分CSV（adult/general）→ Excel集計ファイル自動生成

実行方法:
  python main.py                    # 前日分を自動処理
  python main.py --date 20250615    # 指定日処理
  python main.py --debug            # デバッグモード
"""

import sys
from pathlib import Path
from datetime import datetime, timedelta
import typer
from loguru import logger

# プロジェクトルートをPythonパスに追加
sys.path.insert(0, str(Path(__file__).parent))

from orchestrator import CampaignReportOrchestrator

app = typer.Typer(help="fam8キャンペーンレポート自動集計システム")

@app.command()
def main(
    date: str = typer.Option(
        None, 
        "--date", 
        help="処理対象日 (YYYYMMDD形式, 未指定時は前日)"
    ),
    debug: bool = typer.Option(
        False, 
        "--debug", 
        help="デバッグモード"
    )
):
    """fam8キャンペーンレポート自動集計処理を実行"""
    orchestrator = CampaignReportOrchestrator(debug_mode=debug)
    orchestrator.execute(target_date=date)

if __name__ == "__main__":
    app()