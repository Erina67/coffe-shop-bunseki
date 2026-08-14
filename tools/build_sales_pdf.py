#!/usr/bin/env python3
"""
月次売上分析レポート PDF生成パイプライン

使い方:
    python3 build_sales_pdf.py --excel <build_sales_report.py が出力したxlsx>

仕様の詳細は docs/売上分析レポートPDF_出力仕様.md を参照。
このスクリプトには特定の月名・年月・目標金額・店舗名・商品名・施策名を一切
ハードコードしていない。すべて --excel で渡されるファイル（build_sales_report.py
が生成した「サマリー/日次推移/店舗別/チャネル別/カテゴリ別/顧客区分別」シート）
から読み取る。入力のExcelファイルは一切書き換えない（読み取り専用で開く）。
"""
import argparse
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.font_manager as fm
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.patches import Rectangle
import openpyxl

# ---------------------------------------------------------------------------
# 固定レイアウト定数（データ値ではなく体裁の定数のみ）
# ---------------------------------------------------------------------------
PAGE_W_MM, PAGE_H_MM = 297.0, 210.0  # A4横
MARGIN_MM = 15.0
MM = 1 / 25.4  # inch per mm
PAGE_W_IN, PAGE_H_IN = PAGE_W_MM * MM, PAGE_H_MM * MM

NAVY = "#1F3864"
ALT_ROW = "#F2F2F2"
WHITE = "#FFFFFF"
CHART_COLOR = "#4472C4"  # Excel既定テーマ(Office)のアクセント1と同じ青

FONT_DIR_CANDIDATES = [
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
]
BOLD_FONT_CANDIDATES = [
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
]

SIZE_BODY = 10
SIZE_HEADING = 14
SIZE_PAGE_TITLE = 18

PAGE_TITLES = [
    "1. サマリー",
    "2. 日次推移",
    "3. 店舗別",
    "4. チャネル別",
    "5. カテゴリ別",
    "6. 顧客区分別",
]
DIM_SHEETS = ["店舗別", "チャネル別", "カテゴリ別", "顧客区分別"]


def _find_font(candidates, label):
    for c in candidates:
        if Path(c).exists():
            return c
    raise FileNotFoundError(
        f"{label}用のフォントが見つかりません。`apt-get install fonts-noto-cjk` を実行してください。"
    )


REGULAR_FONT_PATH = _find_font(FONT_DIR_CANDIDATES, "本文")
BOLD_FONT_PATH = _find_font(BOLD_FONT_CANDIDATES, "太字")
FP_REG = fm.FontProperties(fname=REGULAR_FONT_PATH)
FP_BOLD = fm.FontProperties(fname=BOLD_FONT_PATH)


# ---------------------------------------------------------------------------
# 表示用フォーマット（Excel側の書式 ¥#,##0 / 0.0% / yyyy/m/d と同じ見え方にする。
#  数値そのものは一切再計算しない。Excel側が文字列（「記載なし」等）ならそのまま表示）
# ---------------------------------------------------------------------------
def fmt_yen(v):
    if isinstance(v, (int, float)):
        return f"¥{v:,.0f}"
    return "" if v is None else str(v)


def fmt_pct(v):
    if isinstance(v, (int, float)):
        return f"{v * 100:.1f}%"
    return "" if v is None else str(v)


def fmt_int(v):
    if isinstance(v, (int, float)):
        return f"{v:,.0f}"
    return "" if v is None else str(v)


def fmt_date(d):
    return f"{d.year}/{d.month}/{d.day}"


def fmt_plain(v):
    return "" if v is None else str(v)


# ---------------------------------------------------------------------------
# Excelファイルの読み取り（読み取り専用。書き換え・再計算は一切行わない）
# ---------------------------------------------------------------------------
def read_report_data(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)

    sm = wb["サマリー"]
    title = sm.cell(row=1, column=1).value
    kpi_labels = [sm.cell(row=3, column=c).value for c in range(2, 8)]
    kpi_values = [sm.cell(row=4, column=c).value for c in range(2, 8)]
    target = sm.cell(row=6, column=2).value
    campaign_name = sm.cell(row=7, column=2).value
    campaign_period = sm.cell(row=8, column=2).value
    new_items_note = sm.cell(row=10, column=2).value

    daily = wb["日次推移"]
    daily_rows = list(daily.iter_rows(values_only=True))
    daily_header, daily_body = daily_rows[0], daily_rows[1:]
    daily_data = [r for r in daily_body if r[0] != "合計"]
    daily_total = next((r for r in daily_body if r[0] == "合計"), None)

    report_year = daily_data[0][0].year
    report_month = daily_data[0][0].month
    period_start = daily_data[0][0]
    period_end = daily_data[-1][0]

    dim_data = {}
    for name in DIM_SHEETS:
        ws = wb[name]
        rows = list(ws.iter_rows(values_only=True))
        header, body = rows[0], rows[1:]
        items = [r for r in body if r[0] != "合計"]
        total = next((r for r in body if r[0] == "合計"), None)
        dim_data[name] = {"header": header, "items": items, "total": total}

    wb.close()

    return {
        "title": title,
        "kpi_labels": kpi_labels,
        "kpi_values": kpi_values,
        "target": target,
        "campaign_name": campaign_name,
        "campaign_period": campaign_period,
        "new_items_note": new_items_note,
        "report_year": report_year,
        "report_month": report_month,
        "period_start": period_start,
        "period_end": period_end,
        "daily_header": daily_header,
        "daily_data": daily_data,
        "daily_total": daily_total,
        "dim_data": dim_data,
    }


# ---------------------------------------------------------------------------
# ページ共通のヘッダー・フッター
# ---------------------------------------------------------------------------
def draw_header_footer(fig, data, source_filename, page_num, total_pages):
    left_frac = MARGIN_MM / PAGE_W_MM
    right_frac = 1 - MARGIN_MM / PAGE_W_MM
    top_y = 1 - (MARGIN_MM * 0.35) / PAGE_H_MM
    bottom_y = (MARGIN_MM * 0.35) / PAGE_H_MM

    header_left = f"{data['report_year']}年{data['report_month']}月 売上分析レポート"
    header_right = f"対象期間：{fmt_date(data['period_start'])} - {fmt_date(data['period_end'])}"
    fig.text(left_frac, top_y, header_left, fontproperties=FP_BOLD,
              fontsize=SIZE_BODY, ha="left", va="center", color=NAVY)
    fig.text(right_frac, top_y, header_right, fontproperties=FP_REG,
              fontsize=SIZE_BODY, ha="right", va="center", color="black")
    fig.add_artist(plt.Line2D([left_frac, right_frac], [top_y - 0.012, top_y - 0.012],
                               transform=fig.transFigure, color=NAVY, linewidth=0.8))

    fig.text(0.5, bottom_y, f"{page_num} / {total_pages}", fontproperties=FP_REG,
              fontsize=SIZE_BODY, ha="center", va="center", color="black")
    fig.text(right_frac, bottom_y, f"出典：{source_filename}", fontproperties=FP_REG,
              fontsize=SIZE_BODY, ha="right", va="center", color="black")
    fig.add_artist(plt.Line2D([left_frac, right_frac], [bottom_y + 0.012, bottom_y + 0.012],
                               transform=fig.transFigure, color=NAVY, linewidth=0.8))


def new_page_figure(page_title):
    fig = plt.figure(figsize=(PAGE_W_IN, PAGE_H_IN))
    left_frac = MARGIN_MM / PAGE_W_MM
    top_y = 1 - (MARGIN_MM * 0.9) / PAGE_H_MM
    fig.text(left_frac, top_y, page_title, fontproperties=FP_BOLD,
              fontsize=SIZE_PAGE_TITLE, ha="left", va="top", color=NAVY)
    return fig


# ---------------------------------------------------------------------------
# 表（ヘッダー行=濃紺・白・太字、データ行=白/薄グレー交互、横罫のみ）
# ---------------------------------------------------------------------------
def draw_table(fig, rect, headers, rows, col_specs, total_row=None):
    """rect: (left, bottom, width, height) figure-fraction座標
    col_specs: [(header_label, formatter, align), ...] rows と同じ列数
    total_row があれば最終行として太字で描く"""
    left, bottom, width, height = rect
    ax = fig.add_axes([left, bottom, width, height])
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")

    n_data_rows = len(rows) + (1 if total_row is not None else 0)
    n_rows = n_data_rows + 1  # +ヘッダー
    row_h = 1.0 / n_rows

    col_widths = [w for _, _, _, w in col_specs]
    col_lefts = [sum(col_widths[:i]) for i in range(len(col_widths))]

    def y_top(i):
        return 1.0 - i * row_h

    # ヘッダー行
    ax.add_patch(Rectangle((0, y_top(1)), 1, row_h, facecolor=NAVY, edgecolor="none"))
    for (label, _fmt, align, w), cl in zip(col_specs, col_lefts):
        x = cl + (0.01 if align == "left" else w - 0.01)
        ax.text(x, y_top(0) - row_h / 2, label, fontproperties=FP_BOLD,
                 fontsize=SIZE_BODY, color=WHITE, ha=align, va="center")

    all_rows = list(rows)
    for idx, row in enumerate(all_rows, start=1):
        is_total = False
        y0 = y_top(idx + 1)
        if idx % 2 == 0:
            ax.add_patch(Rectangle((0, y0), 1, row_h, facecolor=ALT_ROW, edgecolor="none"))
        for (label, fmt, align, w), cl, val in zip(col_specs, col_lefts, row):
            x = cl + (0.01 if align == "left" else w - 0.01)
            ax.text(x, y0 + row_h / 2, fmt(val), fontproperties=FP_REG,
                     fontsize=SIZE_BODY, color="black", ha=align, va="center")
        ax.add_line(plt.Line2D([0, 1], [y0, y0], color="#BFBFBF", linewidth=0.6))

    if total_row is not None:
        idx = len(all_rows) + 1
        y0 = y_top(idx + 1)
        for col_i, ((label, fmt, align, w), cl, val) in enumerate(zip(col_specs, col_lefts, total_row)):
            x = cl + (0.01 if align == "left" else w - 0.01)
            display_fmt = fmt_plain if col_i == 0 else fmt
            ax.text(x, y0 + row_h / 2, display_fmt(val), fontproperties=FP_BOLD,
                     fontsize=SIZE_BODY, color="black", ha=align, va="center")
        ax.add_line(plt.Line2D([0, 1], [y0, y0], color=NAVY, linewidth=1.2))

    # 表全体の外枠（上端・最下端の横罫）
    ax.add_line(plt.Line2D([0, 1], [1, 1], color=NAVY, linewidth=1.2))
    ax.add_line(plt.Line2D([0, 1], [0, 0], color=NAVY, linewidth=1.2))
    return ax


# ---------------------------------------------------------------------------
# グラフ（Excelと同じ色 #4472C4 を使用。日次推移=折れ線、○○別=棒）
# ---------------------------------------------------------------------------
def _style_chart_axes(ax, xlabel, ylabel):
    ax.set_xlabel(xlabel, fontproperties=FP_REG, fontsize=SIZE_BODY)
    ax.set_ylabel(ylabel, fontproperties=FP_REG, fontsize=SIZE_BODY)
    for lbl in ax.get_xticklabels() + ax.get_yticklabels():
        lbl.set_fontproperties(FP_REG)
        lbl.set_fontsize(SIZE_BODY - 1)
    ax.yaxis.set_major_formatter(lambda v, pos: f"¥{v:,.0f}")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(axis="y", color="#DDDDDD", linewidth=0.6)
    ax.set_axisbelow(True)


def draw_line_chart(fig, rect, dates, values, title):
    ax = fig.add_axes(rect)
    ax.plot(dates, values, color=CHART_COLOR, linewidth=1.6, marker="o", markersize=2.5)
    ax.set_title(title, fontproperties=FP_BOLD, fontsize=SIZE_HEADING, color=NAVY)
    step = max(1, len(dates) // 15)
    ax.set_xticks(dates[::step])
    ax.set_xticklabels([fmt_date(d) for d in dates[::step]], rotation=45, ha="right")
    _style_chart_axes(ax, "日付", "売上金額（円）")
    return ax


def draw_bar_chart(fig, rect, labels, values, title):
    ax = fig.add_axes(rect)
    ax.bar(labels, values, color=CHART_COLOR)
    ax.set_title(title, fontproperties=FP_BOLD, fontsize=SIZE_HEADING, color=NAVY)
    ax.set_xticks(range(len(labels)))
    ax.set_xticklabels(labels, ha="center")
    _style_chart_axes(ax, "", "売上金額（円）")
    return ax


# ---------------------------------------------------------------------------
# ページ構築
# ---------------------------------------------------------------------------
def build_summary_page(pdf, data, source_filename, page_num, total_pages):
    fig = new_page_figure(PAGE_TITLES[0])

    left = MARGIN_MM / PAGE_W_MM
    width = 1 - 2 * MARGIN_MM / PAGE_W_MM
    cur_top = 1 - (MARGIN_MM + 26) / PAGE_H_MM  # ページタイトルの下に十分な余白を取る

    # --- KPI 6指標 ---
    kpi_h = 20 / PAGE_H_MM
    kpi_bottom = cur_top - kpi_h
    n = len(data["kpi_labels"])
    col_w = width / n
    formatters = [fmt_yen, fmt_yen, fmt_pct, fmt_int, fmt_yen, fmt_pct]
    for i, (label, value, f) in enumerate(zip(data["kpi_labels"], data["kpi_values"], formatters)):
        x = left + i * col_w
        ax = fig.add_axes([x, kpi_bottom, col_w * 0.92, kpi_h])
        ax.axis("off")
        ax.set_xlim(0, 1)
        ax.set_ylim(0, 1)
        ax.add_patch(Rectangle((0, 0.55), 1, 0.45, facecolor=NAVY, edgecolor="none"))
        ax.text(0.5, 0.775, label, fontproperties=FP_BOLD, fontsize=SIZE_BODY,
                 color=WHITE, ha="center", va="center")
        display = f(value) if not isinstance(value, str) else value
        ax.text(0.5, 0.25, display, fontproperties=FP_BOLD, fontsize=SIZE_HEADING,
                 color="black", ha="center", va="center")
    cur_top = kpi_bottom - 10 / PAGE_H_MM

    # --- 目標・施策情報 ---
    info_h = 18 / PAGE_H_MM
    info_bottom = cur_top - info_h
    ax_info = fig.add_axes([left, info_bottom, width, info_h])
    ax_info.axis("off")
    ax_info.set_xlim(0, 1)
    ax_info.set_ylim(0, 1)
    info_labels = ["月間売上目標", "販促施策名", "施策対象期間"]
    info_values = [fmt_yen(data["target"]) if isinstance(data["target"], (int, float)) else data["target"],
                    data["campaign_name"], data["campaign_period"]]
    for i, (lbl, val) in enumerate(zip(info_labels, info_values)):
        y = 1 - (i + 0.5) / 3
        ax_info.text(0.0, y, lbl, fontproperties=FP_BOLD, fontsize=SIZE_BODY, ha="left", va="center")
        ax_info.text(0.14, y, str(val), fontproperties=FP_REG, fontsize=SIZE_BODY, ha="left", va="center")
    cur_top = info_bottom - 6 / PAGE_H_MM

    # --- 新規項目 ---
    note_h = 10 / PAGE_H_MM
    note_bottom = cur_top - note_h
    ax_note = fig.add_axes([left, note_bottom, width, note_h])
    ax_note.axis("off")
    ax_note.set_xlim(0, 1)
    ax_note.set_ylim(0, 1)
    ax_note.text(0.0, 0.9, "新規項目", fontproperties=FP_BOLD, fontsize=SIZE_BODY, ha="left", va="top")
    ax_note.text(0.14, 0.9, str(data["new_items_note"]), fontproperties=FP_REG, fontsize=SIZE_BODY,
                  ha="left", va="top", wrap=True)
    cur_top = note_bottom - 10 / PAGE_H_MM

    # --- グラフ（日次推移の折れ線・店舗別の棒） ---
    chart_bottom = MARGIN_MM / PAGE_H_MM + 8 / PAGE_H_MM
    chart_h = cur_top - chart_bottom
    half_w = (width - 10 / PAGE_W_MM) / 2
    dates = [r[0] for r in data["daily_data"]]
    revenues = [r[2] for r in data["daily_data"]]
    draw_line_chart(fig, [left, chart_bottom, half_w, chart_h], dates, revenues, "日次売上推移")

    store = data["dim_data"]["店舗別"]
    labels = [r[0] for r in store["items"]]
    values = [r[2] for r in store["items"]]
    draw_bar_chart(fig, [left + half_w + 10 / PAGE_W_MM, chart_bottom, half_w, chart_h],
                    labels, values, "店舗別売上")

    draw_header_footer(fig, data, source_filename, page_num, total_pages)
    pdf.savefig(fig)
    plt.close(fig)


def build_daily_page(pdf, data, source_filename, page_num, total_pages):
    fig = new_page_figure(PAGE_TITLES[1])
    left = MARGIN_MM / PAGE_W_MM
    width = 1 - 2 * MARGIN_MM / PAGE_W_MM
    top = 1 - (MARGIN_MM + 26) / PAGE_H_MM

    dates = [r[0] for r in data["daily_data"]]
    revenues = [r[2] for r in data["daily_data"]]
    chart_h = 44 / PAGE_H_MM
    draw_line_chart(fig, [left, top - chart_h, width, chart_h], dates, revenues, "日次売上推移")

    table_top = top - chart_h - 22 / PAGE_H_MM
    table_h = table_top - MARGIN_MM / PAGE_H_MM - 8 / PAGE_H_MM
    col_specs = [
        ("日付", fmt_date, "left", 0.26),
        ("件数", fmt_int, "right", 0.13),
        ("売上金額", fmt_yen, "right", 0.25),
        ("粗利", fmt_yen, "right", 0.21),
        ("施策期間", fmt_plain, "left", 0.15),
    ]
    rows = data["daily_data"]
    total = data["daily_total"]
    total_display = (total[0], total[1], total[2], total[3], "") if total else None

    # 縦に31行入り切らないため、月前半/後半の2列に分けて配置する
    mid = (len(rows) + 1) // 2
    left_rows, right_rows = rows[:mid], rows[mid:]
    gap = 8 / PAGE_W_MM
    col_w = (width - gap) / 2
    draw_table(fig, (left, table_top - table_h, col_w, table_h), None, left_rows, col_specs, None)
    draw_table(fig, (left + col_w + gap, table_top - table_h, col_w, table_h), None,
               right_rows, col_specs, total_display)

    draw_header_footer(fig, data, source_filename, page_num, total_pages)
    pdf.savefig(fig)
    plt.close(fig)


def build_dim_page(pdf, data, sheet_name, page_title, source_filename, page_num, total_pages):
    fig = new_page_figure(page_title)
    left = MARGIN_MM / PAGE_W_MM
    width = 1 - 2 * MARGIN_MM / PAGE_W_MM
    top = 1 - (MARGIN_MM + 26) / PAGE_H_MM

    dim = data["dim_data"][sheet_name]
    labels = [r[0] for r in dim["items"]]
    values = [r[2] for r in dim["items"]]
    chart_h = 58 / PAGE_H_MM
    draw_bar_chart(fig, [left, top - chart_h, width, chart_h], labels, values, f"{sheet_name}売上")

    table_top = top - chart_h - 10 / PAGE_H_MM
    table_h = table_top - MARGIN_MM / PAGE_H_MM - 8 / PAGE_H_MM
    col_specs = [
        ("項目", fmt_plain, "left", 0.24),
        ("件数", fmt_int, "right", 0.14),
        ("売上金額", fmt_yen, "right", 0.18),
        ("粗利", fmt_yen, "right", 0.18),
        ("粗利率", fmt_pct, "right", 0.12),
        ("新規項目", fmt_plain, "left", 0.14),
    ]
    rows = dim["items"]
    total = dim["total"]
    total_display = (total[0], total[1], total[2], total[3], total[4], "") if total else None
    draw_table(fig, (left, table_top - table_h, width, table_h), None, rows, col_specs, total_display)

    draw_header_footer(fig, data, source_filename, page_num, total_pages)
    pdf.savefig(fig)
    plt.close(fig)


# ---------------------------------------------------------------------------
# メイン処理
# ---------------------------------------------------------------------------
def build_pdf(excel_path, out_path=None):
    excel_path = Path(excel_path)
    data = read_report_data(excel_path)
    source_filename = excel_path.name

    if out_path is None:
        out_path = excel_path.parent / f"{data['report_year']}年{data['report_month']}月_売上分析レポート.pdf"
    out_path = Path(out_path)

    total_pages = 2 + len(DIM_SHEETS)  # サマリー + 日次推移 + 4つの内訳

    with PdfPages(out_path) as pdf:
        pdf_meta = pdf.infodict()
        pdf_meta["Title"] = f"{data['report_year']}年{data['report_month']}月 売上分析レポート"

        build_summary_page(pdf, data, source_filename, 1, total_pages)
        build_daily_page(pdf, data, source_filename, 2, total_pages)
        for i, sheet_name in enumerate(DIM_SHEETS):
            build_dim_page(pdf, data, sheet_name, PAGE_TITLES[2 + i], source_filename, 3 + i, total_pages)

    return out_path


def main():
    parser = argparse.ArgumentParser(description="月次売上分析レポート PDF生成")
    parser.add_argument("--excel", required=True,
                         help="build_sales_report.py が出力したxlsxファイルのパス（このファイルは書き換えない）")
    parser.add_argument("--out", help="出力PDFのパス（省略時は --excel と同じ場所に "
                                       "『YYYY年M月_売上分析レポート.pdf』として保存）")
    args = parser.parse_args()

    out_path = build_pdf(args.excel, args.out)
    print(f"saved: {out_path}")


if __name__ == "__main__":
    main()
