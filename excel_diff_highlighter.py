"""
Excelファイルの差分比較ツール
新旧バージョンのExcelファイルを比較し、差分がある文字のみを青色でハイライトします。
"""

import openpyxl
from openpyxl.cell.text import InlineFont
from openpyxl.cell.rich_text import TextBlock, CellRichText
import difflib
import re
import time
from pathlib import Path
from typing import Tuple, List, Optional

# 設定
MAX_CELL_VALUE_LENGTH = 100  # 差分サマリーシートに表示するセル値の最大文字数
DEFAULT_HIGHLIGHT_COLOR = '000000FF'  # デフォルトのハイライト色（青、aRGB形式）

# ファイル処理設定
OUTPUT_FILE_SUFFIX = "_差分ハイライト"  # 出力ファイル名のサフィックス
TEMP_FILE_PREFIX = '~$'  # 一時ファイルのプレフィックス
PROGRESS_DISPLAY_INTERVAL = 10  # 進捗表示間隔（%）

# UI設定
SEPARATOR_LENGTH = 60  # 区切り線の長さ

# ハイライト色マップ（aRGB形式）
COLOR_MAP = {
    '1': ('000000FF', '青'),
    '2': ('0000FF00', '緑'),
    '3': ('00FF8C00', 'オレンジ'),
    '4': ('00800080', '紫'),
    '5': ('00FF69B4', 'ピンク'),
    '6': ('00FF0000', '赤')
}

# 差分サマリーシート設定
SUMMARY_SHEET_NAME = "差分サマリー"  # 差分サマリーシート名
SUMMARY_HEADER_COLOR = 'D3D3D3'  # ヘッダー背景色（ライトグレー）
SUMMARY_HEADER_FONT_SIZE = 11  # ヘッダーフォントサイズ
SUMMARY_COL_WIDTH_NO = 8       # No.列の幅
SUMMARY_COL_WIDTH_SHEET = 25   # シート名列の幅
SUMMARY_COL_WIDTH_CELL = 10    # セル列の幅
SUMMARY_COL_WIDTH_VALUE = 40   # 旧値/新値列の幅


def find_file_by_pattern(directory: str, pattern: str) -> List[Path]:
    """
    ファイル名のパターンに基づいてファイルを検索
    例: 「V5_帳票スケッチ_帳票部品No.107_転出名簿4」で検索
    """
    directory_path = Path(directory)
    if not directory_path.exists():
        print(f"エラー: ディレクトリが見つかりません: {directory}")
        return []

    # パターンをエスケープして正規表現として使用
    escaped_pattern = re.escape(pattern)
    # バージョン番号部分を柔軟にマッチ
    regex_pattern = escaped_pattern + r'.*\.xlsx?$'

    matching_files = []
    for file in directory_path.glob('*.xlsx'):
        if re.search(regex_pattern, file.name, re.IGNORECASE):
            matching_files.append(file)

    return matching_files


def extract_version_number(filename: str) -> float:
    """
    ファイル名からバージョン番号を抽出
    例: "v2.06" -> 2.06
    """
    match = re.search(r'[vV](\d+\.\d+)', filename)
    if match:
        return float(match.group(1))
    return 0.0


def extract_base_filename(filename: str) -> str:
    """
    ファイル名からバージョン番号、コピー表記、拡張子を除いた基本名を抽出
    例: "V5_帳票スケッチ_帳票部品No.107_転出名簿4_v2.06.xlsx" -> "V5_帳票スケッチ_帳票部品No.107_転出名簿4"
    例: "V5_帳票スケッチ_帳票部品No.105_転出名簿2_v2.09 のコピー.xlsx" -> "V5_帳票スケッチ_帳票部品No.105_転出名簿2"
    例: "【サイト管理】検索条件・導入元紐づけ整理 のコピー.xlsx" -> "【サイト管理】検索条件・導入元紐づけ整理"
    例: "ファイル名 (1).xlsx" -> "ファイル名"
    """
    # 拡張子を除去
    name_without_ext = Path(filename).stem

    # バージョン番号部分とその後の文字列（のコピー、など）を除去
    # [_\s]*: アンダースコアまたは空白文字（複数可）
    # [vV]\d+\.\d+: バージョン番号（v2.06など）
    # .*$: バージョン番号以降のすべての文字（" のコピー"など）
    base_name = re.sub(r'[_\s]*[vV]\d+\.\d+.*$', '', name_without_ext)

    # バージョン番号がない場合でも「のコピー」「(1)」などを除去
    # \s*: 空白文字（複数可）
    # (のコピー|\(\d+\)|copy): 「のコピー」「(数字)」「copy」などのパターン
    base_name = re.sub(r'\s*(のコピー|\(\d+\)|copy|\s-\s*コピー).*$', '', base_name, flags=re.IGNORECASE)

    return base_name.strip()


def find_matching_file_pairs(old_directory: str, new_directory: str) -> Tuple[List[Tuple[str, str, str]], List[str], List[str]]:
    """
    新旧ディレクトリから対応するファイルペアを検索
    戻り値: (pairs, unmatched_old_files, unmatched_new_files)
        pairs: [(base_name, old_file_path, new_file_path), ...]
        unmatched_old_files: 旧フォルダにのみ存在するファイル名のリスト
        unmatched_new_files: 新フォルダにのみ存在するファイル名のリスト
    """
    old_dir = Path(old_directory)
    new_dir = Path(new_directory)

    if not old_dir.exists():
        print(f"エラー: 旧ディレクトリが見つかりません: {old_directory}")
        return ([], [], [])

    if not new_dir.exists():
        print(f"エラー: 新ディレクトリが見つかりません: {new_directory}")
        return ([], [], [])

    # 旧ディレクトリのファイルを基本名でグループ化
    old_files = {}
    for file in old_dir.glob('*.xlsx'):
        if file.name.startswith(TEMP_FILE_PREFIX):  # 一時ファイルをスキップ
            continue
        base_name = extract_base_filename(file.name)
        if base_name not in old_files:
            old_files[base_name] = []
        old_files[base_name].append(file)

    # 新ディレクトリのファイルを基本名でグループ化
    new_files = {}
    for file in new_dir.glob('*.xlsx'):
        if file.name.startswith(TEMP_FILE_PREFIX):  # 一時ファイルをスキップ
            continue
        base_name = extract_base_filename(file.name)
        if base_name not in new_files:
            new_files[base_name] = []
        new_files[base_name].append(file)

    # マッチングするペアを検索
    pairs = []
    matched_bases = set()

    for base_name in old_files:
        if base_name in new_files:
            # 各グループ内でファイルを選択
            # バージョン番号がある場合は最新を選択、ない場合は最初のファイルを選択
            old_versions = [(f, extract_version_number(f.name)) for f in old_files[base_name]]
            new_versions = [(f, extract_version_number(f.name)) for f in new_files[base_name]]

            # バージョン番号が存在する場合（0.0より大きい）は最新を選択
            if any(v > 0 for _, v in old_versions):
                old_file = max(old_files[base_name], key=lambda f: extract_version_number(f.name))
            else:
                # バージョン番号がない場合は最初のファイル
                old_file = old_files[base_name][0]

            if any(v > 0 for _, v in new_versions):
                new_file = max(new_files[base_name], key=lambda f: extract_version_number(f.name))
            else:
                # バージョン番号がない場合は最初のファイル
                new_file = new_files[base_name][0]

            # 同じファイルでないことを確認（パスが異なる場合は処理）
            if str(old_file) != str(new_file):
                pairs.append((base_name, str(old_file), str(new_file)))
                matched_bases.add(base_name)

    # マッチングしないファイルを報告
    unmatched_old = set(old_files.keys()) - matched_bases
    unmatched_new = set(new_files.keys()) - matched_bases

    # マッチングしなかったファイル名をリストに格納
    unmatched_old_files = []
    unmatched_new_files = []

    if unmatched_old:
        print(f"\n⚠ 旧フォルダにのみ存在するファイル（新バージョンなし）:")
        for base_name in sorted(unmatched_old):
            for file in old_files[base_name]:
                print(f"  - {file.name}")
                unmatched_old_files.append(file.name)

    if unmatched_new:
        print(f"\n⚠ 新フォルダにのみ存在するファイル（旧バージョンなし）:")
        for base_name in sorted(unmatched_new):
            for file in new_files[base_name]:
                print(f"  - {file.name}")
                unmatched_new_files.append(file.name)

    return pairs, unmatched_old_files, unmatched_new_files


def find_old_and_new_versions(directory: str, base_filename: str) -> Tuple[Optional[str], Optional[str]]:
    """
    指定されたディレクトリから新旧のバージョンファイルを検索
    """
    files = find_file_by_pattern(directory, base_filename)

    if not files or len(files) < 2:
        print(f"エラー: {base_filename} に一致するファイルが2つ以上見つかりません")
        if files:
            print(f"見つかったファイル: {[f.name for f in files]}")
        return None, None

    # バージョン番号でソート
    sorted_files = sorted(files, key=lambda f: extract_version_number(f.name))

    old_file = str(sorted_files[-2])  # 2番目に新しいファイル（古いバージョン）
    new_file = str(sorted_files[-1])  # 最新ファイル（新しいバージョン）

    print(f"古いバージョン: {Path(old_file).name}")
    print(f"新しいバージョン: {Path(new_file).name}")

    return old_file, new_file


def get_cell_value_as_string(cell) -> str:
    """
    セルの値を文字列として取得

    Args:
        cell: 対象セル
    """
    if cell.value is None:
        return ""
    return str(cell.value)


def find_char_differences(old_text: str, new_text: str) -> List[Tuple[int, int]]:
    """
    2つのテキスト間の文字レベルの差分を検出
    戻り値: [(start_index, end_index), ...] 差分がある文字の範囲リスト
    """
    if old_text == new_text:
        return []

    # 文字レベルでの差分を検出
    matcher = difflib.SequenceMatcher(None, old_text, new_text)
    differences = []

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag in ('replace', 'insert'):
            # 新しいテキスト側の変更された部分
            differences.append((j1, j2))

    return differences


def apply_blue_color_to_differences(cell, old_text: str, new_text: str, highlight_color: str = DEFAULT_HIGHLIGHT_COLOR):
    """
    差分がある文字のみを指定色にする

    Args:
        cell: 対象セル
        old_text: 旧テキスト
        new_text: 新テキスト
        highlight_color: ハイライト色（aRGB形式の16進数）
    """
    differences = find_char_differences(old_text, new_text)

    if not differences:
        return

    # 元のセルのフォント情報を取得
    original_font = cell.font

    # RichTextオブジェクトを作成
    rich_text_parts = []
    current_pos = 0

    # InlineFontのパラメータを準備（サポートされているもののみ）
    normal_font_kwargs = {}
    blue_font_kwargs = {}

    # フォントサイズ
    if original_font.size:
        normal_font_kwargs['sz'] = original_font.size
        blue_font_kwargs['sz'] = original_font.size

    # フォント名
    if original_font.name:
        normal_font_kwargs['rFont'] = original_font.name
        blue_font_kwargs['rFont'] = original_font.name

    # 元の色（aRGB形式の8文字16進数に変換）
    if original_font.color and original_font.color.rgb:
        try:
            color_value = str(original_font.color.rgb)
            # 16進数カラーコードの検証と変換
            color_value = color_value.upper().strip()
            # 英数字のみを抽出
            color_value = ''.join(c for c in color_value if c in '0123456789ABCDEF')

            if len(color_value) == 6:
                # RGB形式の場合、先頭に'00'（アルファチャンネル）を追加
                color_value = '00' + color_value
            elif len(color_value) == 8:
                # すでにaRGB形式
                pass
            else:
                # 不正な形式の場合は色を設定しない
                color_value = None

            if color_value and len(color_value) == 8:
                normal_font_kwargs['color'] = color_value
        except Exception:
            # 色の変換に失敗した場合はスキップ
            pass

    # ハイライト色
    blue_font_kwargs['color'] = highlight_color

    # 下線
    if original_font.underline:
        normal_font_kwargs['u'] = original_font.underline
        blue_font_kwargs['u'] = original_font.underline

    # フォントオブジェクトを作成
    normal_font = InlineFont(**{k: v for k, v in normal_font_kwargs.items() if v is not None})
    blue_font = InlineFont(**{k: v for k, v in blue_font_kwargs.items() if v is not None})

    for start, end in differences:
        # 差分の前の通常テキスト
        if current_pos < start:
            rich_text_parts.append(TextBlock(normal_font, new_text[current_pos:start]))

        # 差分部分（青色）
        if start < end:
            rich_text_parts.append(TextBlock(blue_font, new_text[start:end]))

        current_pos = end

    # 残りのテキスト
    if current_pos < len(new_text):
        rich_text_parts.append(TextBlock(normal_font, new_text[current_pos:]))

    # セルにRichTextを設定
    if rich_text_parts:
        cell.value = CellRichText(*rich_text_parts)


def compare_and_highlight_excel(old_file_path: str, new_file_path: str, output_file_path: str, highlight_color: str = DEFAULT_HIGHLIGHT_COLOR, compare_formulas: bool = False):
    """
    2つのExcelファイルを比較し、差分を指定色でハイライト

    Args:
        old_file_path: 旧ファイルのパス
        new_file_path: 新ファイルのパス
        output_file_path: 出力ファイルのパス
        highlight_color: ハイライト色（aRGB形式の16進数、デフォルトは青）
        compare_formulas: Trueの場合は数式を比較、Falseの場合は表示値を比較
    """
    print(f"\n処理開始...")
    print(f"古いファイル: {old_file_path}")
    print(f"新しいファイル: {new_file_path}")
    print(f"比較モード: {'数式' if compare_formulas else '表示値'}")

    # 処理開始時刻を記録
    start_time = time.time()

    # ファイルを開く（compare_formulasがTrueの場合は数式を保持、Falseの場合は表示値のみ）
    try:
        old_wb = openpyxl.load_workbook(old_file_path, data_only=not compare_formulas)
    except Exception as e:
        print(f"エラー: 旧ファイルを開けませんでした: {e}")
        raise

    try:
        new_wb = openpyxl.load_workbook(new_file_path, data_only=not compare_formulas)
    except Exception as e:
        print(f"エラー: 新ファイルを開けませんでした: {e}")
        old_wb.close()
        raise

    changes_log = []  # 変更履歴を記録

    # 全シートを比較
    for sheet_name in new_wb.sheetnames:
        if sheet_name not in old_wb.sheetnames:
            print(f"警告: シート '{sheet_name}' は古いファイルに存在しません")
            continue

        old_sheet = old_wb[sheet_name]
        new_sheet = new_wb[sheet_name]

        print(f"\nシート '{sheet_name}' を処理中...")
        sheet_changes = 0

        # 総セル数を計算
        total_cells = new_sheet.max_row * new_sheet.max_column
        processed_cells = 0
        last_progress = 0

        # 各セルを比較
        for row in range(1, new_sheet.max_row + 1):
            for col in range(1, new_sheet.max_column + 1):
                old_cell = old_sheet.cell(row, col)
                new_cell = new_sheet.cell(row, col)

                old_value = get_cell_value_as_string(old_cell)
                new_value = get_cell_value_as_string(new_cell)

                # 両方空セルの場合はスキップ（パフォーマンス向上）
                if not old_value and not new_value:
                    processed_cells += 1
                    continue

                # 差分がある場合
                if old_value != new_value and new_value:
                    apply_blue_color_to_differences(new_cell, old_value, new_value, highlight_color)
                    sheet_changes += 1

                    # 変更履歴を記録
                    changes_log.append({
                        'sheet': sheet_name,
                        'cell': f'{new_cell.column_letter}{new_cell.row}',
                        'old': old_value[:MAX_CELL_VALUE_LENGTH] + ('...' if len(old_value) > MAX_CELL_VALUE_LENGTH else ''),
                        'new': new_value[:MAX_CELL_VALUE_LENGTH] + ('...' if len(new_value) > MAX_CELL_VALUE_LENGTH else '')
                    })

                # 進行状況を表示（10%刻み）
                processed_cells += 1
                progress = int((processed_cells / total_cells) * 100)
                if progress >= last_progress + PROGRESS_DISPLAY_INTERVAL and progress < 100:
                    print(f"  進行状況: {progress}% ({processed_cells}/{total_cells} セル)")
                    last_progress = progress
    # 差分サマリーシートを作成
    if changes_log:
        print(f"\n差分サマリーシートを作成中...")
        summary_sheet = new_wb.create_sheet(SUMMARY_SHEET_NAME, 0)  # 最初のシートとして追加

        # ヘッダー行を追加
        summary_sheet['A1'] = 'No.'
        summary_sheet['B1'] = 'シート名'
        summary_sheet['C1'] = 'セル'
        summary_sheet['D1'] = '旧値'
        summary_sheet['E1'] = '新値'

        # ヘッダーのスタイル設定
        from openpyxl.styles import Font, PatternFill, Alignment
        header_font = Font(bold=True, size=SUMMARY_HEADER_FONT_SIZE)
        header_fill = PatternFill(start_color=SUMMARY_HEADER_COLOR, end_color=SUMMARY_HEADER_COLOR, fill_type='solid')
        header_alignment = Alignment(horizontal='center', vertical='center')

        for cell in summary_sheet[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment

        # 変更履歴を書き込み
        for idx, change in enumerate(changes_log, start=2):
            summary_sheet[f'A{idx}'] = idx - 1
            summary_sheet[f'B{idx}'] = change['sheet']
            summary_sheet[f'C{idx}'] = change['cell']
            summary_sheet[f'D{idx}'] = change['old']
            summary_sheet[f'E{idx}'] = change['new']

        # 列幅を調整
        summary_sheet.column_dimensions['A'].width = SUMMARY_COL_WIDTH_NO
        summary_sheet.column_dimensions['B'].width = SUMMARY_COL_WIDTH_SHEET
        summary_sheet.column_dimensions['C'].width = SUMMARY_COL_WIDTH_CELL
        summary_sheet.column_dimensions['D'].width = SUMMARY_COL_WIDTH_VALUE
        summary_sheet.column_dimensions['E'].width = SUMMARY_COL_WIDTH_VALUE

        print(f"  差分サマリーシートに {len(changes_log)} 件の変更を記録しました")

    # 結果を保存
    new_wb.save(output_file_path)

    # 処理時間を計算
    elapsed_time = time.time() - start_time

    # 結果表示
    if len(changes_log) == 0:
        print(f"\n完了！ 差分は見つかりませんでした（新旧ファイルは同一です）")
    else:
        print(f"\n完了！ 合計 {len(changes_log)} 個のセルに差分が見つかりました")
    print(f"処理時間: {elapsed_time:.1f}秒")
    print(f"出力ファイル: {output_file_path}")

    old_wb.close()
    new_wb.close()


def main():
    """
    メイン処理
    """
    print("=" * SEPARATOR_LENGTH)
    print("Excel差分ハイライトツール")
    print("=" * SEPARATOR_LENGTH)
    print("\n旧バージョンと新バージョンがそれぞれ別のフォルダにあり、")
    print("同じファイル名（ベース名）のペアをすべて自動処理します。")

    # ハイライト色の選択
    print("\n差分をハイライトする色を選択してください:")
    for key, (_, color_name) in sorted(COLOR_MAP.items()):
        print(f"{key}. {color_name}")

    color_choice = input(f"選択 (1-{len(COLOR_MAP)}, デフォルト: 1): ").strip()

    if color_choice not in COLOR_MAP:
        color_choice = '1'  # デフォルトは青

    highlight_color, color_name = COLOR_MAP[color_choice]
    print(f"選択された色: {color_name}\n")

    # 比較モードの選択
    print("差分比較のモードを選択してください:")
    print("1. 表示値のみ比較（数式は比較しない）")
    print("2. 数式を比較（数式がある場合は数式を比較）")

    mode_choice = input("選択 (1-2, デフォルト: 1): ").strip()
    compare_formulas = (mode_choice == '2')

    mode_name = "数式" if compare_formulas else "表示値"
    print(f"選択されたモード: {mode_name}\n")

    # フォルダ配下の全ファイルを一括処理
    print("旧バージョンのフォルダを指定してください")
    old_directory = input("旧バージョンのフォルダパス: ").strip().strip('"').strip("'")
    if not old_directory:
        old_directory = "."

    print("\n新バージョンのフォルダを指定してください")
    new_directory = input("新バージョンのフォルダパス: ").strip().strip('"').strip("'")
    if not new_directory:
        new_directory = "."

    print("\n出力先フォルダを指定してください")
    output_directory = input("出力先フォルダパス（空欄で新バージョンと同じ）: ").strip().strip('"').strip("'")
    if not output_directory:
        output_directory = new_directory

    # マッチングするファイルペアを検索
    file_pairs, unmatched_old_files, unmatched_new_files = find_matching_file_pairs(old_directory, new_directory)

    if not file_pairs:
        print("\nマッチングするファイルペアが見つかりませんでした")
        return

    print(f"\n{len(file_pairs)} 個のファイルペアが見つかりました:")
    for i, (base_name, old_file, new_file) in enumerate(file_pairs, 1):
        print(f"{i}. {base_name}")
        print(f"   旧: {Path(old_file).name}")
        print(f"   新: {Path(new_file).name}")

    if not file_pairs:
        print("\nファイルが見つからなかったため処理を終了します")
        return

    # 出力ディレクトリを作成（存在しない場合）
    output_path = Path(output_directory)
    if not output_path.exists():
        output_path.mkdir(parents=True, exist_ok=True)

    # 確認
    if len(file_pairs) == 1:
        base_name, old_file, new_file = file_pairs[0]
        new_file_path = Path(new_file)
        output_filename = new_file_path.stem + OUTPUT_FILE_SUFFIX + new_file_path.suffix
        output_file = str(output_path / output_filename)
        print(f"\n出力ファイル: {output_filename}")
    else:
        print(f"\n出力先: {output_directory}")
        print(f"処理対象: {len(file_pairs)} ファイル")

    confirm = input("\n処理を開始しますか？ (y/n): ").strip().lower()

    if confirm != 'y':
        print("処理をキャンセルしました")
        return

    # 比較とハイライト処理
    success_count = 0
    error_count = 0

    for i, (base_name, old_file, new_file) in enumerate(file_pairs, 1):
        try:
            print(f"\n{'='*SEPARATOR_LENGTH}")
            print(f"[{i}/{len(file_pairs)}] 処理中: {Path(new_file).name}")
            print(f"{'='*SEPARATOR_LENGTH}")

            new_file_path = Path(new_file)
            output_filename = new_file_path.stem + OUTPUT_FILE_SUFFIX + new_file_path.suffix
            output_file = str(output_path / output_filename)

            compare_and_highlight_excel(old_file, new_file, output_file, highlight_color, compare_formulas)
            success_count += 1

        except Exception as e:
            print(f"\nエラーが発生しました: {e}")
            error_count += 1
            import traceback
            traceback.print_exc()

    # 最終結果
    print(f"\n{'='*SEPARATOR_LENGTH}")
    print(f"処理完了")
    print(f"{'='*SEPARATOR_LENGTH}")
    print(f"成功: {success_count} ファイル")
    print(f"失敗: {error_count} ファイル")
    print(f"出力先: {output_directory}")

    # マッチングしなかったファイルの報告
    if unmatched_old_files:
        print(f"\n旧フォルダにのみ存在（{len(unmatched_old_files)} ファイル）:")
        for filename in unmatched_old_files:
            print(f"  - {filename}")

    if unmatched_new_files:
        print(f"\n新フォルダにのみ存在（{len(unmatched_new_files)} ファイル）:")
        for filename in unmatched_new_files:
            print(f"  - {filename}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n処理を中断しました")
    except Exception as e:
        print(f"\n\n予期しないエラーが発生しました: {e}")
        import traceback
        traceback.print_exc()
    finally:
        input("\nEnterキーを押して終了...")
