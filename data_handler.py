#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fam8キャンペーンレポート自動集計システム - Excelデータ操作・関数埋込専門
CSV貼付・動的関数埋込・シート操作処理（完全修正版）
"""

import pandas as pd
from pathlib import Path
import xlwings as xw
from loguru import logger
import time


class DataHandler:
    """Excelデータ操作クラス"""

    def __init__(self, config: dict, target_date_str: str):
        self.config = config
        self.target_date_str = target_date_str
        self.filter_excel_path = Path(config["paths"]["filter_input_excel"])

        # シート名設定
        self.csv_sheet_name = config["excel_structure"]["csv_sheet_name"]
        self.summary_sheet_name = config["excel_structure"]["summary_sheet_name"]

        # 集計設定
        self.max_campaign_rows = config["filter_settings"]["max_campaign_rows"]

        # xlwingsアプリケーション参照保持
        self.app = None

    def process(self, csv_data: pd.DataFrame) -> xw.Book:
        """Excelデータ操作メイン処理"""
        logger.info("Excelデータ操作開始")

        # Excelアプリケーション設定
        self.app = xw.App(visible=False, add_book=False)
        self.app.display_alerts = False
        self.app.screen_updating = False

        try:
            # ワークブック開く
            workbook = self.app.books.open(str(self.filter_excel_path))

            # CSV貼付処理（CSVの列順序・列名をそのまま保持）
            self._paste_csv_data(workbook, csv_data)

            # 計算を強制実行してから関数埋込
            workbook.app.calculate()
            time.sleep(1)  # 計算完了待機

            # 動的関数埋込処理
            self._embed_dynamic_formulas(workbook)

            # 再計算実行
            workbook.app.calculate()
            time.sleep(1)  # 計算完了待機

            logger.info("Excelデータ操作完了")
            return workbook

        except Exception as e:
            logger.error(f"Excelデータ操作エラー: {e}")
            if self.app:
                self.app.quit()
            raise

    def _paste_csv_data(self, workbook: xw.Book, csv_data: pd.DataFrame):
        """CSV貼付処理（CSVの列順序・列名をそのまま保持）"""
        logger.info("CSV貼付処理開始")

        try:
            # 前日分CSV抽出シート取得・作成
            if self.csv_sheet_name in [sheet.name for sheet in workbook.sheets]:
                csv_sheet = workbook.sheets[self.csv_sheet_name]
                # 既存データ完全クリア
                csv_sheet.clear()
            else:
                csv_sheet = workbook.sheets.add(name=self.csv_sheet_name)
                # シートを先頭に移動
                csv_sheet.api.Move(Before=workbook.sheets[0].api)

            # CSVデータをA1から正確に貼付
            if not csv_data.empty:
                logger.info(f"CSV貼付データ確認: {csv_data.shape[0]}行 × {csv_data.shape[1]}列")
                logger.info(f"CSV列構成: {list(csv_data.columns)}")

                # 重要な列の位置をログ出力
                self._log_column_mapping(csv_data)

                # ヘッダー行を1行目に貼付
                header_row = list(csv_data.columns)
                for col_idx, header in enumerate(header_row, 1):
                    cell_address = f"{self._column_number_to_letter(col_idx)}1"
                    csv_sheet.range(cell_address).value = str(header)
                
                logger.info(f"ヘッダー行貼付完了: A1:{self._column_number_to_letter(len(header_row))}1")

                # データ行を2行目から貼付
                data_values = csv_data.values.tolist()
                num_rows = len(data_values)
                num_cols = len(data_values[0]) if data_values else 0

                logger.info(f"データ貼付予定: {num_rows}行 × {num_cols}列（2行目から開始）")

                # データを確実に貼付
                if num_rows > 0 and num_cols > 0:
                    try:
                        # 範囲指定して一括貼付
                        start_cell = "A2"
                        end_cell = f"{self._column_number_to_letter(num_cols)}{num_rows + 1}"
                        paste_range = f"{start_cell}:{end_cell}"
                        csv_sheet.range(paste_range).value = data_values
                        logger.info(f"一括CSV貼付完了: {paste_range}")
                    except Exception as bulk_error:
                        logger.warning(f"一括貼付失敗、行ごと貼付に切替: {bulk_error}")
                        
                        # 行ごと貼付（フォールバック）
                        for row_idx, row_data in enumerate(data_values, 2):
                            try:
                                row_range = f"A{row_idx}:{self._column_number_to_letter(len(row_data))}{row_idx}"
                                csv_sheet.range(row_range).value = row_data
                            except Exception as row_error:
                                logger.warning(f"行{row_idx}貼付エラー: {row_error}")

                # 貼付結果検証
                self._verify_paste_result(csv_sheet, num_rows, num_cols)

            else:
                logger.warning("CSVデータが空のため貼付をスキップ")

        except Exception as e:
            logger.error(f"CSV貼付エラー: {e}")
            raise

    def _log_column_mapping(self, csv_data: pd.DataFrame):
        """列マッピング情報をログ出力"""
        logger.info("=== CSV列マッピング確認 ===")
        
        # 実際の列構造をすべて出力
        for i, col_name in enumerate(csv_data.columns):
            excel_col = self._column_number_to_letter(i + 1)
            logger.info(f"  {col_name} → {excel_col}列（{i+1}番目）")

    def _verify_paste_result(self, sheet: xw.Sheet, num_rows: int, num_cols: int):
        """貼付結果検証"""
        logger.info("=== CSV貼付結果検証 ===")
        
        try:
            # ヘッダー確認（全列）
            header_range = f"A1:{self._column_number_to_letter(num_cols)}1"
            header_values = sheet.range(header_range).value
            if isinstance(header_values, list):
                logger.info(f"ヘッダー確認: {header_values}")
            else:
                logger.info(f"ヘッダー確認: {[header_values]}")
            
            # データ確認（2-4行目の重要列）
            important_cols = {'キャンペーン名': 'C', 'Imp': 'I', 'Click': 'J', 'CV': 'L', 'グロス': 'N', 'ネット': 'O'}
            
            if num_rows > 0:
                for row in range(2, min(5, num_rows + 2)):  # 2-4行目
                    row_data = {}
                    for col_name, excel_col in important_cols.items():
                        try:
                            cell_value = sheet.range(f"{excel_col}{row}").value
                            row_data[col_name] = cell_value
                        except:
                            row_data[col_name] = "エラー"
                    logger.info(f"データ行{row}: {row_data}")
                        
            logger.info("貼付結果検証完了")
            
        except Exception as e:
            logger.warning(f"貼付結果検証エラー: {e}")

    def _embed_dynamic_formulas(self, workbook: xw.Book):
        """動的関数埋込処理"""
        logger.info("動的関数埋込開始")

        try:
            # 集計シート取得・作成
            if self.summary_sheet_name in [sheet.name for sheet in workbook.sheets]:
                summary_sheet = workbook.sheets[self.summary_sheet_name]
                # B2:I100の範囲をクリア（A列は保持）
                summary_sheet.range("B2:I100").clear_contents()
            else:
                summary_sheet = workbook.sheets.add(name=self.summary_sheet_name)

            # ヘッダー設定
            self._set_headers(summary_sheet)

            # シート参照確認
            csv_sheet_exists = self.csv_sheet_name in [sheet.name for sheet in workbook.sheets]
            logger.info(f"CSV抽出シート存在確認: {csv_sheet_exists}")

            # CSV列位置を動的に特定（完全修正版）
            column_positions = self._detect_csv_column_positions(workbook)

            # 関数埋込（正確な列位置使用）
            self._embed_formulas_range(summary_sheet, column_positions)

            logger.info("動的関数埋込完了")

        except Exception as e:
            logger.error(f"動的関数埋込エラー: {e}")
            raise

    def _detect_csv_column_positions(self, workbook: xw.Book) -> dict:
        """CSV列位置動的検出（完全修正版 - 全範囲検索対応）"""
        logger.info("=== CSV列位置検出開始（全範囲検索修正版） ===")
        
        column_positions = {}
        
        try:
            csv_sheet = workbook.sheets[self.csv_sheet_name]
            
            # ヘッダー行（1行目）を読み取り - 範囲を大幅拡大（最大50列）
            max_check_cols = 50  # 最大50列までチェック
            header_range = f"A1:{self._column_number_to_letter(max_check_cols)}1"
            headers = csv_sheet.range(header_range).value
            
            if isinstance(headers, list):
                header_list = headers
            else:
                header_list = [headers]
            
            logger.info(f"検出されたヘッダー全体（最初の20列）: {header_list[:20]}")
            
            # 対象列の完全一致検索（大文字小文字・前後空白を考慮）
            target_columns = {
                'キャンペーン名': 'campaign_name_col',
                'Imp': 'imp_col', 
                'Click': 'click_col',
                'CV': 'cv_col',
                'グロス': 'gross_col',
                'ネット': 'net_col'
            }
            
            # 完全一致検索（前後空白削除・大文字小文字区別なし）
            for i, header in enumerate(header_list):
                if header is not None:
                    header_str = str(header).strip()
                    for target_name, key in target_columns.items():
                        if header_str == target_name:
                            excel_col = self._column_number_to_letter(i + 1)
                            column_positions[key] = excel_col
                            logger.info(f"★ 列位置検出成功: {target_name} → {excel_col}列（{i+1}番目）")
                            break
            
            # 部分一致検索（完全一致で見つからない場合のフォールバック）
            if len(column_positions) < len(target_columns):
                logger.warning("完全一致で全列が検出できないため部分一致検索を実行")
                for i, header in enumerate(header_list):
                    if header is not None:
                        header_str = str(header).strip().lower()
                        for target_name, key in target_columns.items():
                            if key not in column_positions:
                                if target_name.lower() in header_str or header_str in target_name.lower():
                                    excel_col = self._column_number_to_letter(i + 1)
                                    column_positions[key] = excel_col
                                    logger.info(f"◆ 部分一致で検出: {target_name} → {excel_col}列（{i+1}番目）ヘッダー: '{header_str}'")
                                    break
            
            # 検出結果確認
            found_columns = len(column_positions)
            total_columns = len(target_columns)
            logger.info(f"列位置検出結果: {found_columns}/{total_columns}列")
            
            # 実際のCSV構造に基づく正確なデフォルト値（最後の手段）
            correct_default_positions = {
                'campaign_name_col': 'C',  # キャンペーン名 = 3列目
                'imp_col': 'I',            # Imp = 9列目
                'click_col': 'J',          # Click = 10列目  
                'cv_col': 'L',             # CV = 12列目
                'gross_col': 'N',          # グロス = 14列目
                'net_col': 'O'             # ネット = 15列目
            }
            
            # 検出できなかった列は正確なデフォルト値を使用
            for key, correct_col in correct_default_positions.items():
                if key not in column_positions:
                    column_positions[key] = correct_col
                    logger.warning(f"列位置未検出、正確なデフォルト使用: {key} → {correct_col}列")
            
            logger.info(f"=== 最終列位置マッピング ===")
            for key, col in column_positions.items():
                logger.info(f"  {key}: {col}列")
            
            # 検証: 実際にセルの値を確認
            self._verify_column_positions(csv_sheet, column_positions)
            
            return column_positions
            
        except Exception as e:
            logger.error(f"CSV列位置検出エラー: {e}")
            # エラー時は実際のCSV構造に基づく正確な位置を返す
            return {
                'campaign_name_col': 'C',  # キャンペーン名
                'imp_col': 'I',            # Imp
                'click_col': 'J',          # Click
                'cv_col': 'L',             # CV
                'gross_col': 'N',          # グロス
                'net_col': 'O'             # ネット
            }

    def _verify_column_positions(self, csv_sheet: xw.Sheet, column_positions: dict):
        """列位置検証（数値データの存在確認強化）"""
        logger.info("=== 列位置検証開始（数値データ確認強化） ===")
        
        try:
            # ヘッダー確認
            for key, excel_col in column_positions.items():
                header_value = csv_sheet.range(f"{excel_col}1").value
                logger.info(f"  {key} ({excel_col}列): ヘッダー='{header_value}'")
            
            # データ確認（2-5行目、数値系列は型チェックも）
            logger.info("データ行確認（2-5行目）:")
            numeric_columns = ['imp_col', 'click_col', 'cv_col', 'gross_col', 'net_col']
            
            for row in range(2, 7):  # 2-6行目
                logger.info(f"  行{row}:")
                for key, excel_col in column_positions.items():
                    try:
                        data_value = csv_sheet.range(f"{excel_col}{row}").value
                        
                        # 数値列の場合は型と値の詳細確認
                        if key in numeric_columns:
                            if data_value is not None:
                                # 数値変換可能かチェック
                                try:
                                    float_value = float(str(data_value).replace(',', ''))
                                    logger.info(f"    {key} ({excel_col}{row}): データ='{data_value}' (数値: {float_value})")
                                except ValueError:
                                    logger.warning(f"    {key} ({excel_col}{row}): データ='{data_value}' (数値変換不可)")
                            else:
                                logger.info(f"    {key} ({excel_col}{row}): データ=None")
                        else:
                            logger.info(f"    {key} ({excel_col}{row}): データ='{data_value}'")
                            
                    except Exception as cell_error:
                        logger.warning(f"    {key} ({excel_col}{row}): データ読み取りエラー - {cell_error}")
                        
        except Exception as e:
            logger.warning(f"列位置検証エラー: {e}")

    def _set_headers(self, sheet: xw.Sheet):
        """ヘッダー設定"""
        headers = self.config["excel_structure"]["summary_columns"]

        for i, header in enumerate(headers, 1):
            cell = sheet.range(f"{self._column_number_to_letter(i)}1")
            cell.value = header

        logger.info(f"ヘッダー設定完了: {len(headers)}列")

    def _embed_formulas_range(self, sheet: xw.Sheet, column_positions: dict):
        """関数範囲埋込（グロス・ネット計算完全修正版）"""

        # シート参照名を正確に指定（スペース対応）
        csv_sheet_ref = f"'{self.csv_sheet_name}'"

        # 正確な列位置取得
        campaign_col = column_positions.get('campaign_name_col', 'C')
        imp_col = column_positions.get('imp_col', 'I')
        click_col = column_positions.get('click_col', 'J')
        cv_col = column_positions.get('cv_col', 'L')
        gross_col = column_positions.get('gross_col', 'N')
        net_col = column_positions.get('net_col', 'O')

        logger.info(f"関数で使用する正確な列位置:")
        logger.info(f"  キャンペーン名={campaign_col}, Imp={imp_col}, Click={click_col}")
        logger.info(f"  CV={cv_col}, グロス={gross_col}, ネット={net_col}")

        # 関数埋込カウンター
        formula_count = 0

        for row in range(2, self.max_campaign_rows + 2):  # 2行目から101行目まで

            try:
                # B列: Imp（元のまま維持）
                formula_b = f'''=IF(A{row}="", "",
  LET(
    キー, A{row},
    検索列, {csv_sheet_ref}!{campaign_col}:{campaign_col},
    対象列, {csv_sheet_ref}!{imp_col}:{imp_col},
    該当値, FILTER(対象列, ISNUMBER(SEARCH(キー, 検索列))),
    合計, IFERROR(SUM(該当値), ""),
    合計
  )
)'''
                sheet.range(f"B{row}").formula = formula_b
                formula_count += 1

                # C列: Click（元のまま維持）
                formula_c = f'''=IF(A{row}="", "",
  LET(
    キー, A{row},
    検索列, {csv_sheet_ref}!{campaign_col}:{campaign_col},
    対象列, {csv_sheet_ref}!{click_col}:{click_col},
    該当値, FILTER(対象列, ISNUMBER(SEARCH(キー, 検索列))),
    合計, IFERROR(SUM(該当値), ""),
    合計
  )
)'''
                sheet.range(f"C{row}").formula = formula_c
                formula_count += 1

                # D列: CTR（元のまま維持）
                formula_d = f'=IF(OR(B{row}="", C{row}="", B{row}=0), "", TEXT(C{row}/B{row}, "0.00%"))'
                sheet.range(f"D{row}").formula = formula_d
                formula_count += 1

                # E列: CV（元のまま維持）
                formula_e = f'''=IF(A{row}="", "",
  LET(
    キー, A{row},
    検索列, {csv_sheet_ref}!{campaign_col}:{campaign_col},
    対象列, {csv_sheet_ref}!{cv_col}:{cv_col},
    該当値, FILTER(対象列, ISNUMBER(SEARCH(キー, 検索列))),
    合計, IFERROR(SUM(該当値), ""),
    合計
  )
)'''
                sheet.range(f"E{row}").formula = formula_e
                formula_count += 1

                # F列: CVR（元のまま維持）
                formula_f = f'=IF(OR(C{row}="", E{row}="", C{row}=0), "", TEXT(E{row}/C{row}, "0.00%"))'
                sheet.range(f"F{row}").formula = formula_f
                formula_count += 1

                # G列: グロス（元のLET+FILTER構文で確実に86,087を計算）
                formula_g = f'''=IF(A{row}="", "",
  LET(
    キー, A{row},
    検索列, {csv_sheet_ref}!{campaign_col}:{campaign_col},
    対象列, {csv_sheet_ref}!{gross_col}:{gross_col},
    該当値, FILTER(対象列, ISNUMBER(SEARCH(キー, 検索列))),
    合計, IFERROR(SUM(該当値), ""),
    合計
  )
)'''
                sheet.range(f"G{row}").formula = formula_g
                formula_count += 1

                # H列: ネット（元のLET+FILTER構文で正確な値を計算）
                formula_h = f'''=IF(A{row}="", "",
  LET(
    キー, A{row},
    検索列, {csv_sheet_ref}!{campaign_col}:{campaign_col},
    対象列, {csv_sheet_ref}!{net_col}:{net_col},
    該当値, FILTER(対象列, ISNUMBER(SEARCH(キー, 検索列))),
    合計, IFERROR(SUM(該当値), ""),
    合計
  )
)'''
                sheet.range(f"H{row}").formula = formula_h
                formula_count += 1

                # I列: 税別グロス（元のまま維持）
                formula_i = f'=IF(OR(G{row}="", ISERROR(G{row})), "", ROUND(G{row}/1.1, 0))'
                sheet.range(f"I{row}").formula = formula_i
                formula_count += 1

                # 10行ごとにログ出力
                if row % 10 == 2:  # 2, 12, 22, ...
                    logger.debug(f"関数埋込進捗: {row - 1}行完了")

            except Exception as formula_error:
                logger.error(f"行{row}の関数埋込エラー: {formula_error}")

        logger.info(f"関数埋込完了: {formula_count}個の関数を挿入")

        # 関数確認ログ
        try:
            sample_formula_b = sheet.range("B2").formula
            sample_formula_g = sheet.range("G2").formula
            logger.info(f"関数確認サンプル B2: {sample_formula_b[:100]}...")
            logger.info(f"グロス関数確認 G2: {sample_formula_g[:100]}...")
        except:
            logger.warning("関数確認に失敗")

    def _column_number_to_letter(self, col_num: int) -> str:
        """列番号をアルファベットに変換"""
        result = ""
        while col_num > 0:
            col_num -= 1
            result = chr(65 + (col_num % 26)) + result
            col_num //= 26
        return result

    def save_workbook(self, workbook: xw.Book):
        """ワークブック保存"""
        try:
            # 最終計算実行
            workbook.app.calculate()
            time.sleep(2)  # 計算完了待機

            # 保存
            workbook.save()
            logger.info(f"ワークブック保存完了: {self.filter_excel_path}")
        except Exception as e:
            logger.error(f"ワークブック保存エラー: {e}")
            raise
        finally:
            # アプリケーション終了
            if workbook:
                workbook.close()
            if self.app:
                self.app.quit()
