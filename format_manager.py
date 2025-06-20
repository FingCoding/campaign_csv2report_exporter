#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fam8キャンペーンレポート自動集計システム - Excel書式設定・見た目制御専門
CTR/CVR右寄せ・ヘッダーグレー・グリッド線・数値フォーマット設定（修正版）
"""

import xlwings as xw
from loguru import logger


class FormatManager:
    """Excel書式設定クラス"""
    
    def __init__(self, config: dict):
        self.config = config
        
        # 書式設定
        self.summary_sheet_name = config["excel_structure"]["summary_sheet_name"]
        self.max_campaign_rows = config["filter_settings"]["max_campaign_rows"]
        
        # 書式定義
        self.percentage_format = config["excel_formatting"]["percentage_format"]
        self.currency_format = config["excel_formatting"]["currency_format"]
        self.number_format = config["excel_formatting"]["number_format"]
        self.header_background_color = tuple(config["excel_formatting"]["header_background_color"])
    
    def apply_formatting(self, workbook: xw.Book):
        """書式設定メイン処理"""
        logger.info("Excel書式設定開始")
        
        try:
            # 集計シート取得
            summary_sheet = workbook.sheets[self.summary_sheet_name]
            
            # 計算実行（書式設定前に数値を確定）
            workbook.app.calculate()
            
            # ヘッダー書式設定
            self._format_headers(summary_sheet)
            
            # CTR/CVR列右寄せ設定
            self._format_percentage_columns(summary_sheet)
            
            # 数値列書式設定
            self._format_number_columns(summary_sheet)
            
            # 通貨列書式設定
            self._format_currency_columns(summary_sheet)
            
            # 列幅自動調整
            self._auto_adjust_columns(summary_sheet)
            
            # グリッド線設定
            self._apply_grid_lines(summary_sheet)
            
            # 最終計算実行
            workbook.app.calculate()
            
            logger.info("Excel書式設定完了")
            
        except Exception as e:
            logger.error(f"Excel書式設定エラー: {e}")
            raise
    
    def _format_headers(self, sheet: xw.Sheet):
        """ヘッダー書式設定"""
        logger.info("ヘッダー書式設定開始")
        
        try:
            # ヘッダー範囲（A1:I1）
            header_range = sheet.range("A1:I1")
            
            # フォント設定
            header_range.api.Font.Bold = True
            header_range.api.Font.Size = 11
            header_range.api.Font.Name = "メイリオ"
            
            # 背景色設定（薄いグレー）
            header_range.color = self.header_background_color
            
            # 文字色設定（黒）
            header_range.api.Font.Color = 0x000000
            
            # 中央揃え
            header_range.api.HorizontalAlignment = -4108  # xlCenter
            header_range.api.VerticalAlignment = -4108    # xlCenter
            
            # 罫線設定
            try:
                # 外枠罫線
                header_range.api.Borders(7).Weight = 2   # xlEdgeLeft
                header_range.api.Borders(8).Weight = 2   # xlEdgeTop
                header_range.api.Borders(9).Weight = 2   # xlEdgeBottom
                header_range.api.Borders(10).Weight = 2  # xlEdgeRight
                
                # 内側縦罫線
                header_range.api.Borders(11).Weight = 1  # xlInsideVertical
                
                # 罫線色
                for border_id in [7, 8, 9, 10, 11]:
                    try:
                        header_range.api.Borders(border_id).Color = 0x000000
                    except:
                        pass
            except Exception as border_error:
                logger.warning(f"ヘッダー罫線設定エラー: {border_error}")
            
            logger.info("ヘッダー書式設定完了")
            
        except Exception as e:
            logger.error(f"ヘッダー書式設定エラー: {e}")
            raise
    
    def _format_percentage_columns(self, sheet: xw.Sheet):
        """CTR/CVR列右寄せ・パーセント書式設定"""
        logger.info("CTR/CVR列書式設定開始")
        
        try:
            # CTR列（D列）とCVR列（F列）の範囲
            ctr_range = sheet.range(f"D2:D{self.max_campaign_rows + 1}")
            cvr_range = sheet.range(f"F2:F{self.max_campaign_rows + 1}")
            
            # CTR列書式設定
            try:
                ctr_range.api.HorizontalAlignment = -4152  # xlRight（右寄せ）
                # パーセント書式は関数内で TEXT() を使用しているため適用しない
            except Exception as e:
                logger.warning(f"CTR列書式設定エラー: {e}")
            
            # CVR列書式設定
            try:
                cvr_range.api.HorizontalAlignment = -4152  # xlRight（右寄せ）
                # パーセント書式は関数内で TEXT() を使用しているため適用しない
            except Exception as e:
                logger.warning(f"CVR列書式設定エラー: {e}")
            
            logger.info("CTR/CVR列書式設定完了")
            
        except Exception as e:
            logger.error(f"CTR/CVR列書式設定エラー: {e}")
            raise
    
    def _format_number_columns(self, sheet: xw.Sheet):
        """数値列書式設定"""
        logger.info("数値列書式設定開始")
        
        try:
            # 数値列: B(Imp), C(Click), E(CV), G(グロス), H(ネット)
            number_columns = ["B", "C", "E", "G", "H"]
            
            for col in number_columns:
                try:
                    col_range = sheet.range(f"{col}2:{col}{self.max_campaign_rows + 1}")
                    col_range.api.NumberFormat = self.number_format
                    col_range.api.HorizontalAlignment = -4152  # xlRight（右寄せ）
                except Exception as e:
                    logger.warning(f"{col}列書式設定エラー: {e}")
            
            logger.info("数値列書式設定完了")
            
        except Exception as e:
            logger.error(f"数値列書式設定エラー: {e}")
            raise
    
    def _format_currency_columns(self, sheet: xw.Sheet):
        """通貨列書式設定"""
        logger.info("通貨列書式設定開始")
        
        try:
            # 税別グロス列（I列）
            currency_range = sheet.range(f"I2:I{self.max_campaign_rows + 1}")
            currency_range.api.NumberFormat = self.currency_format
            currency_range.api.HorizontalAlignment = -4152  # xlRight（右寄せ）
            
            logger.info("通貨列書式設定完了")
            
        except Exception as e:
            logger.error(f"通貨列書式設定エラー: {e}")
            raise
    
    def _auto_adjust_columns(self, sheet: xw.Sheet):
        """列幅自動調整"""
        logger.info("列幅自動調整開始")
        
        try:
            # A列からI列まで自動調整
            for col in range(1, 10):  # A=1, B=2, ..., I=9
                try:
                    col_letter = self._column_number_to_letter(col)
                    sheet.range(f"{col_letter}:{col_letter}").autofit()
                    
                    # 最小列幅設定（見やすさのため）
                    min_width = 12 if col_letter in ["D", "F"] else 10  # CTR/CVRは少し広く
                    current_width = sheet.range(f"{col_letter}1").column_width
                    if current_width < min_width:
                        sheet.range(f"{col_letter}:{col_letter}").column_width = min_width
                        
                except Exception as e:
                    logger.warning(f"{col_letter}列幅調整エラー: {e}")
            
            logger.info("列幅自動調整完了")
            
        except Exception as e:
            logger.error(f"列幅自動調整エラー: {e}")
            raise
    
    def _apply_grid_lines(self, sheet: xw.Sheet):
        """グリッド線設定"""
        logger.info("グリッド線設定開始")
        
        try:
            # データ範囲特定
            data_range = self._get_data_range(sheet)
            
            if data_range:
                try:
                    # 罫線設定
                    range_obj = sheet.range(data_range)
                    
                    # 外枠罫線（太線）
                    for border_id in [7, 8, 9, 10]:  # Left, Top, Bottom, Right
                        try:
                            range_obj.api.Borders(border_id).Weight = 3
                            range_obj.api.Borders(border_id).Color = 0x000000
                        except:
                            pass
                    
                    # 内側罫線（細線）
                    try:
                        range_obj.api.Borders(11).Weight = 1  # xlInsideVertical
                        range_obj.api.Borders(11).Color = 0x808080
                        range_obj.api.Borders(12).Weight = 1  # xlInsideHorizontal
                        range_obj.api.Borders(12).Color = 0x808080
                    except:
                        pass
                    
                    logger.info(f"グリッド線設定完了: {data_range}")
                    
                except Exception as border_error:
                    logger.warning(f"グリッド線設定エラー: {border_error}")
            else:
                logger.warning("データ範囲が特定できないためグリッド線設定をスキップ")
            
        except Exception as e:
            logger.error(f"グリッド線設定エラー: {e}")
            raise
    
    def _get_data_range(self, sheet: xw.Sheet) -> str:
        """データ範囲特定"""
        try:
            # A列の最終行を検索（キャンペーン名が入っている行まで）
            last_row = 1
            for row in range(2, self.max_campaign_rows + 2):
                try:
                    cell_value = sheet.range(f"A{row}").value
                    if cell_value and str(cell_value).strip():
                        last_row = row
                except:
                    break
            
            if last_row > 1:
                return f"A1:I{last_row}"
            else:
                # データがない場合はヘッダーのみ
                return "A1:I1"
            
        except Exception as e:
            logger.error(f"データ範囲特定エラー: {e}")
            return None
    
    def _column_number_to_letter(self, col_num: int) -> str:
        """列番号をアルファベットに変換"""
        result = ""
        while col_num > 0:
            col_num -= 1
            result = chr(65 + (col_num % 26)) + result
            col_num //= 26
        return result