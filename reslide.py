#!/usr/bin/env python3
"""ReSlide: .pptx を走査し、見やすさの問題をレポートする（フェーズ1）。

使い方: python reslide.py <対象.pptx> [--config 設定.json] [--out レポート.md]
       python reslide.py <対象.pptx> --fix   … 自動調整（フェーズ2）
       python reslide.py <対象.pptx> --split … 分割候補の判定だけをレポートする（フェーズ3）
       python reslide.py <対象.pptx> --fix --split … 自動調整に加えてページ分割も行う
       python reslide.py <対象.pptx> --reorder … 並び替え候補だけをレポートする（フェーズ4）
       python reslide.py <対象.pptx> --reorder-apply <候補番号> … 選んだ候補だけ並び替えて保存する
       python reslide.py <対象.pptx> --structure-apply <番号> … 構成の指摘を選んで修正して保存する
       python reslide.py <対象.pptx> --gui … 構成の指摘を一覧で選んで修正するウィンドウを開く
       python reslide.py <対象.pptx> --check-update … 新しいバージョンがあるかだけを確認する
元の .pptx は一切変更しない（開くだけで保存はしない。実行前後のハッシュも照合する）。
--fix の結果は <元名>_reslide.pptx へ新規保存し、変更一覧を Markdown で出す。
"""

import argparse
import copy
import hashlib
import json
import math
import os
import re
import sys
import tempfile
import threading
import unicodedata
import urllib.error
import urllib.request
import webbrowser
from pathlib import Path

from lxml import etree

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_FILL
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.util import Pt

__version__ = "1.2.1"

# バージョン更新の確認先（GitHub の最新リリース）
UPDATE_API_URL = "https://api.github.com/repos/Akatuki1121/ReSlide/releases/latest"
UPDATE_RELEASES_URL = "https://github.com/Akatuki1121/ReSlide/releases"
UPDATE_USER_AGENT = "ReSlide-UpdateChecker"
UPDATE_TIMEOUT_SEC = 10

EMU_PER_MM = 36000.0
NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}
A = "{%s}" % NS["a"]

# 解決元の名前（レポートに出す）
SRC_RUN = "run(a:rPr)"
SRC_PPR = "段落(a:pPr/a:defRPr)"
SRC_END = "段落末(a:endParaRPr)"
SRC_LST = "枠のlstStyle"
SRC_LAYOUT = "レイアウトのlstStyle"
SRC_MASTER_PH = "マスターのlstStyle"
SRC_MASTER = "マスターのtxStyles"
SRC_DEFAULT = "presentationの既定(defaultTextStyle)"
DEFAULT_SOURCES = {SRC_MASTER, SRC_DEFAULT}


def mm(emu):
    return emu / EMU_PER_MM


# ---------------------------------------------------------------- フォントサイズ解決
def _sz(el):
    """要素の sz 属性(1/100pt)を pt で返す。無ければ None。"""
    if el is None:
        return None
    v = el.get("sz")
    return int(v) / 100.0 if v else None


def _lvl_defrpr(lststyle, lvl):
    if lststyle is None:
        return None
    return lststyle.find("a:lvl%dpPr/a:defRPr" % (lvl + 1), NS)


class Resolver:
    def __init__(self, prs):
        self.prs = prs
        self.default_style = prs.part._element.find("p:defaultTextStyle", NS)

    @staticmethod
    def _ph_kind(shape):
        """'title' / 'body' / 'other' （マスターのどのtxStylesを使うか）"""
        if shape is None or not shape.is_placeholder:
            return "other"
        t = shape.placeholder_format.type
        if t in (
            PP_PLACEHOLDER.TITLE,
            PP_PLACEHOLDER.CENTER_TITLE,
            PP_PLACEHOLDER.VERTICAL_TITLE,
        ):
            return "title"
        if t in (
            PP_PLACEHOLDER.SLIDE_NUMBER,
            PP_PLACEHOLDER.FOOTER,
            PP_PLACEHOLDER.DATE,
        ):
            return "other"
        return "body"

    def _layout_lst(self, shape):
        if shape is None or not shape.is_placeholder:
            return None, None
        slide = shape.part.slide if hasattr(shape.part, "slide") else None
        if slide is None:
            return None, None
        layout = slide.slide_layout
        lph = None
        for ph in layout.placeholders:
            if ph.placeholder_format.idx == shape.placeholder_format.idx:
                lph = ph
                break
        if lph is None:
            for ph in layout.placeholders:
                if ph.placeholder_format.type == shape.placeholder_format.type:
                    lph = ph
                    break
        mph = None
        for ph in layout.slide_master.placeholders:
            if ph.placeholder_format.type == shape.placeholder_format.type or (
                self._ph_kind(ph) == "body" and self._ph_kind(shape) == "body"
            ):
                mph = ph
                break
        f = lambda s: None if s is None else s._element.find("p:txBody/a:lstStyle", NS)
        return f(lph), f(mph)

    def resolve(self, shape, p_el, r_el, lvl_override=None):
        """1つの run（r_el=None なら段落末）のサイズを (pt, 解決元) で返す。未解決は (None, None)。"""
        ppr = p_el.find("a:pPr", NS)
        lvl = int(ppr.get("lvl", 0)) if ppr is not None else 0
        if lvl_override is not None:
            lvl = lvl_override
        if r_el is not None:
            s = _sz(r_el.find("a:rPr", NS))
            if s:
                return s, SRC_RUN
        if ppr is not None:
            s = _sz(ppr.find("a:defRPr", NS))
            if s:
                return s, SRC_PPR
        end = p_el.find("a:endParaRPr", NS)
        if end is not None and r_el is None:
            s = _sz(end)
            if s:
                return s, SRC_END
        # 枠の lstStyle
        if shape is not None and getattr(shape, "_element", None) is not None:
            lst = shape._element.find("p:txBody/a:lstStyle", NS)
            if lst is None:
                lst = shape._element.find("a:lstStyle", NS)
            s = _sz(_lvl_defrpr(lst, lvl))
            if s:
                return s, SRC_LST
        # レイアウト → マスターのプレースホルダ
        try:
            llst, mlst = self._layout_lst(shape)
        except Exception:
            llst, mlst = None, None
        s = _sz(_lvl_defrpr(llst, lvl))
        if s:
            return s, SRC_LAYOUT
        s = _sz(_lvl_defrpr(mlst, lvl))
        if s:
            return s, SRC_MASTER_PH
        # マスターの txStyles
        kind = self._ph_kind(shape)
        try:
            master = shape.part.slide.slide_layout.slide_master
            tx = master._element.find("p:txStyles", NS)
            style = tx.find(
                {
                    "title": "p:titleStyle",
                    "body": "p:bodyStyle",
                    "other": "p:otherStyle",
                }[kind],
                NS,
            )
            s = _sz(_lvl_defrpr(style, lvl))
            if s:
                return s, SRC_MASTER
        except Exception:
            pass
        s = _sz(_lvl_defrpr(self.default_style, lvl))
        if s:
            return s, SRC_DEFAULT
        return None, None


def autofit_scale(shape):
    """normAutofit の fontScale（縮小率）。無ければ 1.0。"""
    try:
        body = shape._element.find("p:txBody/a:bodyPr", NS)
        na = body.find("a:normAutofit", NS) if body is not None else None
        if na is not None and na.get("fontScale"):
            return int(na.get("fontScale")) / 100000.0
    except Exception:
        pass
    return 1.0


# ---------------------------------------------------------------- 読み取り
class TextUnit:
    """テキスト枠1つ、または表のセル1つ分のテキスト。"""

    def __init__(self, label, paragraphs, scale, inner_w=0, inner_h=0):
        self.label = label  # 例: "" / "セル(2,3)"
        self.paragraphs = paragraphs  # [{"text","runs":[(n,size,src)],"run_els":[(要素,文字列)],"colors":[(r,g,b)|None]}]
        self.scale = scale
        self.inner_w = inner_w  # 枠の内側（余白を差し引いた）幅・EMU・絶対座標
        self.inner_h = inner_h  # 同・高さ
        self.bg_state = None  # 自身の塗り（表のセル・図形自身）の状態。None は背景の決め方を後段に委ねる

    @property
    def text(self):
        return "\n".join(p["text"] for p in self.paragraphs if p["text"].strip())


def _srgb_of_props(props_el):
    """rPr 相当の要素に明示された srgbClr を (r, g, b) で返す。無ければ None。

    取るのは明示の RGB だけ。テーマ色（schemeClr 等）や自動色は解決できない
    ので対象にしない。
    """
    if props_el is None:
        return None
    srgb = props_el.find("a:solidFill/a:srgbClr", NS)
    if srgb is None:
        return None
    v = srgb.get("val") or ""
    if len(v) != 6:
        return None
    try:
        return (int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16))
    except ValueError:
        return None


def _run_color(p_el, r_el):
    """run の文字色 (r, g, b)。run の rPr → 段落の defRPr → endParaRPr の順に探す。

    いずれもスライド内に明示された RGB だけを見る。継承（レイアウトや
    マスター）やテーマ色は辿らない。
    """
    color = _srgb_of_props(r_el.find("a:rPr", NS))
    if color is not None:
        return color
    ppr = p_el.find("a:pPr", NS)
    if ppr is not None:
        color = _srgb_of_props(ppr.find("a:defRPr", NS))
        if color is not None:
            return color
    return _srgb_of_props(p_el.find("a:endParaRPr", NS))


def read_text_unit(resolver, shape, txbody, label="", inner_w=0, inner_h=0):
    scale = autofit_scale(shape) if label == "" else 1.0
    paras = []
    for p_el in txbody.findall("a:p", NS):
        runs, run_els, colors, text_parts = [], [], [], []
        for child in p_el:
            if child.tag in (A + "r", A + "fld"):
                t = child.find("a:t", NS)
                txt = (t.text or "") if t is not None else ""
                if txt:
                    size, src = resolver.resolve(shape, p_el, child)
                    runs.append((len(txt.replace("\n", "")), size, src))
                    run_els.append((child, txt))
                    colors.append(_run_color(p_el, child))
                text_parts.append(txt)
            elif child.tag == A + "br":
                text_parts.append("\n")
        paras.append(
            {
                "text": "".join(text_parts),
                "runs": runs,
                "run_els": run_els,
                "colors": colors,
            }
        )
    return TextUnit(label, paras, scale, inner_w, inner_h)


DEFAULT_INSET_LR = 91440  # 0.1 inch（OOXML の既定の左右余白）
DEFAULT_INSET_TB = 45720  # 0.05 inch（同・上下余白）


def _insets(obj):
    """text_frame / cell の内側余白 (左, 右, 上, 下) を EMU で返す。未指定は既定値。"""

    def g(name, default):
        v = getattr(obj, name, None)
        return v if v is not None else default

    return (
        g("margin_left", DEFAULT_INSET_LR),
        g("margin_right", DEFAULT_INSET_LR),
        g("margin_top", DEFAULT_INSET_TB),
        g("margin_bottom", DEFAULT_INSET_TB),
    )


class Elem:
    def __init__(self, slide_no, kind, name, x, y, w, h):
        self.slide_no, self.kind, self.name = slide_no, kind, name
        self.x, self.y, self.w, self.h = x, y, w, h  # EMU・スライド絶対座標
        self.units = []
        self.ph_type = None  # PP_PLACEHOLDER 名
        self.in_group = False
        self.shape = (
            None  # python-pptx の要素オブジェクト（--fix で位置とサイズを触る）
        )
        self.tf = None  # 絶対座標への変換（グループ内要素の移動量の換算に使う）

    @property
    def has_text(self):
        return any(u.text.strip() for u in self.units)

    @property
    def text(self):
        return " / ".join(u.text for u in self.units if u.text.strip())


class Tf:  # 座標変換 abs = child * s + t
    def __init__(self, sx=1.0, sy=1.0, tx=0.0, ty=0.0):
        self.sx, self.sy, self.tx, self.ty = sx, sy, tx, ty

    def pt(self, x, y):
        return x * self.sx + self.tx, y * self.sy + self.ty


def _kind(shape):
    st = shape.shape_type
    if getattr(shape, "has_table", False) and shape.has_table:
        return "表"
    if getattr(shape, "has_chart", False) and shape.has_chart:
        return "グラフ"
    if st == MSO_SHAPE_TYPE.PICTURE:
        return "図(画像)"
    if st == MSO_SHAPE_TYPE.GROUP:
        return "グループ"
    if st == MSO_SHAPE_TYPE.LINE or shape._element.tag.endswith("}cxnSp"):
        return "線・コネクタ"
    if st == MSO_SHAPE_TYPE.PLACEHOLDER:
        return "プレースホルダ"
    if st == MSO_SHAPE_TYPE.TEXT_BOX:
        return "テキスト枠"
    if shape._element.tag.endswith("}graphicFrame"):
        uri = shape._element.find(".//a:graphicData", NS)
        if uri is not None and "diagram" in (uri.get("uri") or ""):
            return "SmartArt"
        return "その他(graphicFrame)"
    return "図形"


def walk(resolver, slide_no, shapes, tf, out, in_group=False):
    for sh in shapes:
        try:
            x, y, w, h = sh.left, sh.top, sh.width, sh.height
        except Exception:
            x = y = w = h = None
        if x is None or w is None:
            continue
        ax, ay = tf.pt(x, y)
        aw, ah = w * tf.sx, h * tf.sy
        kind = _kind(sh)
        if kind == "グループ":
            xf = sh._element.find("p:grpSpPr/a:xfrm", NS)
            ch_off, ch_ext = xf.find("a:chOff", NS), xf.find("a:chExt", NS)
            cox, coy = int(ch_off.get("x")), int(ch_off.get("y"))
            cw, chh = int(ch_ext.get("cx")) or 1, int(ch_ext.get("cy")) or 1
            sx, sy = aw / cw, ah / chh
            sub = Tf(sx, sy, ax - cox * sx, ay - coy * sy)
            walk(resolver, slide_no, sh.shapes, sub, out, True)
            continue
        el = Elem(slide_no, kind, sh.name, ax, ay, aw, ah)
        el.in_group = in_group
        el.shape = sh
        el.tf = tf
        if sh.is_placeholder:
            try:
                el.ph_type = sh.placeholder_format.type.name
            except Exception:
                el.ph_type = None
        if kind == "表":
            tbl = sh.table
            for ri, row in enumerate(tbl.rows):
                for ci, cell in enumerate(row.cells):
                    if cell.is_spanned:
                        continue
                    tb = cell._tc.find("a:txBody", NS)
                    if tb is not None:
                        il, ir, it, ib = _insets(cell)
                        iw = (tbl.columns[ci].width - il - ir) * tf.sx
                        ih = (tbl.rows[ri].height - it - ib) * tf.sy
                        unit = read_text_unit(
                            resolver,
                            sh,
                            tb,
                            "セル(%d行%d列)" % (ri + 1, ci + 1),
                            iw,
                            ih,
                        )
                        # セルの文字の背景は、まずセル自身の塗りで決まる
                        try:
                            cell_state = _fill_format_state(cell.fill)
                        except Exception:
                            cell_state = ("painted", None)
                        if cell_state[0] != "none":
                            unit.bg_state = cell_state
                        el.units.append(unit)
        elif sh.has_text_frame:
            il, ir, it, ib = _insets(sh.text_frame)
            iw = (w - il - ir) * tf.sx
            ih = (h - it - ib) * tf.sy
            unit = read_text_unit(
                resolver, sh, sh._element.find("p:txBody", NS), "", iw, ih
            )
            # 図形自身に塗りがあれば、その文字の背景はその塗りになる
            own_state = _fill_state(el)
            if own_state[0] != "none":
                unit.bg_state = own_state
            el.units.append(unit)
        out.append(el)


def read_presentation(prs, resolver):
    slides = []
    for i, slide in enumerate(prs.slides, 1):
        elems = []
        walk(resolver, i, slide.shapes, Tf(), elems)
        slides.append(elems)
    return slides


# ---------------------------------------------------------------- 判定
def _min_size(unit):
    """ユニット内の解決済みフォントサイズの最小値（実効値）。無ければ None。"""
    sizes = [
        size * unit.scale
        for p in unit.paragraphs
        for (n, size, src) in p["runs"]
        if size is not None
    ]
    return min(sizes) if sizes else None


def _label_shape_chars(unit, cfg):
    """ラベルの形を満たす場合に、そのテキストの文字数を返す。

    形とは、表のセルでなく、1段落のみ・改行なし・設定の文字数以下であること。
    満たさなければ None。
    """
    if unit.label != "":
        return None
    texts = [p["text"] for p in unit.paragraphs if p["text"].strip()]
    if len(texts) != 1 or "\n" in texts[0]:
        return None
    n = len(texts[0].strip())
    if n > cfg.get("label_max_chars", 30):
        return None
    return n


def classify(elem, unit, cfg, body_median=None):
    """(役割, 根拠) を返す。役割: title / note / label / body

    body_median は、そのスライドの本文の中央値（slide_body_median の結果）。
    ラベルの形をした短文でも、大きさが本文の中央値とほぼ同じものは、
    図の名前ではなく要約の一文など本文の役割とみなして本文に返す。
    """
    if elem.ph_type in cfg["title_placeholder_types"]:
        return "title", "タイトル用プレースホルダ(%s)" % elem.ph_type
    if elem.ph_type in cfg["note_placeholder_types"]:
        return "note", "注記用プレースホルダ(%s)" % elem.ph_type
    first = next((p["text"].strip() for p in unit.paragraphs if p["text"].strip()), "")
    for pre in cfg["note_text_prefixes"]:
        if first.startswith(pre):
            return "note", "先頭が「%s」で始まる注記" % pre
    if first and re.match(cfg["page_number_regex"], unit.text.strip()):
        return "note", "ページ番号の書式（数字のみ／n/m）"
    n = _label_shape_chars(unit, cfg)
    if n is not None:
        if body_median is not None:
            size = _min_size(unit)
            if size is not None:
                ratio = size / body_median
                if (
                    cfg.get("relative_min_ratio", 0.7)
                    <= ratio
                    <= cfg.get("label_body_max_ratio", 1.15)
                ):
                    return (
                        "body",
                        "大きさ %.1fpt がスライドの本文の中央値 %.1fpt とほぼ同じため本文"
                        % (size, body_median),
                    )
        return "label", "1段落・改行なし・%d文字" % n
    return "body", "見出し・注記・ラベルに当たらない本文"


def slide_body_median(elems, cfg):
    """そのスライドの本文の中央値（文字数で重み付け）を返す。無ければ None。

    ラベルの形をした短文は本文に数えない。それ以外の本文が2つ以上ある
    スライドだけを対象にする（1つだけでは「周り」が無いため）。
    """
    pairs = []
    n_body = 0
    for e in elems:
        for u in e.units:
            if not u.text.strip():
                continue
            role, _why = classify(e, u, cfg)
            if role != "body":
                continue
            n_body += 1
            for p in u.paragraphs:
                for n, size, _src in p["runs"]:
                    if size is not None:
                        pairs.append((size * u.scale, n))
    if n_body < 2:
        return None
    return _weighted_median(pairs)


ROLE_JP = {"body": "本文", "label": "ラベル"}


def _weighted_median(pairs):
    """(サイズ, 文字数) の組から、文字数で重み付けした中央値を返す。無ければ None。

    本文の中央値は文字数で重み付ける。短いセル1つ分の文字と長文の本文を
    同じ重さで数えると、表のあるスライドで中央値が実態より小さく出るため。
    """
    total = sum(n for _size, n in pairs)
    if not total:
        return None
    acc = 0
    for size, n in sorted(pairs):
        acc += n
        if acc >= total / 2:
            return size
    return None


def _luminance(rgb):
    """相対輝度（WCAG の定義）。"""

    def f(c):
        c = c / 255.0
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4

    r, g, b = rgb
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)


def _contrast_ratio(c1, c2):
    """2色のコントラスト比（WCAG の定義）。1.0〜21.0。"""
    l1, l2 = _luminance(c1), _luminance(c2)
    hi, lo = max(l1, l2), min(l1, l2)
    return (hi + 0.05) / (lo + 0.05)


def _fill_format_state(fill):
    """塗り（FillFormat）の状態を ("rgb", 色) / ("painted", None) / ("none", None) で返す。

    "rgb" は単色の塗りで RGB が取れたもの。"painted" は色は付いているが RGB で
    決められないもの（グラデーション・パターン・テーマ色の単色など）。
    "none" は塗りが無い（または透明な）もの。
    """
    try:
        t = fill.type
    except Exception:
        return ("painted", None)
    if t == MSO_FILL.SOLID:
        try:
            fc = fill.fore_color
            if fc.type == MSO_COLOR_TYPE.RGB:
                rgb = fc.rgb
                return ("rgb", (int(rgb[0]), int(rgb[1]), int(rgb[2])))
        except Exception:
            pass
        return ("painted", None)
    if t is None or t == MSO_FILL.BACKGROUND:
        return ("none", None)
    return ("painted", None)


def _fill_state(elem):
    """要素の塗りの状態を _fill_format_state と同じ形式で返す。

    画像は中身の色が決められないため "painted"。表・グラフ・SmartArt は塗りを
    1色で表せないため、塗りの取得に失敗したら "painted" とする。
    """
    if elem.kind == "図(画像)":
        return ("painted", None)
    if elem.shape is None:
        return ("none", None)
    try:
        fill = elem.shape.fill
    except Exception:
        return ("painted", None)
    return _fill_format_state(fill)


def _theme_scheme_colors(master):
    """マスターのテーマの色定義（clrScheme）を {名前: (r, g, b)} で返す。

    srgbClr の値と、sysClr の lastClr（実際に表示される色）だけを解決する。
    取れない色は含めない。bg1 や tx1 などの別名は、マスターの clrMap が
    示す対応（bg1 → lt1 など）で引けるようにして返す。clrMap が無ければ
    既定の対応を使う。
    """
    out = {}
    try:
        theme_part = None
        for rel in master.part.rels.values():
            if rel.reltype.endswith("/theme"):
                theme_part = rel.target_part
                break
        if theme_part is None:
            return out
        root = etree.fromstring(theme_part.blob)
        scheme = root.find("a:themeElements/a:clrScheme", NS)
        if scheme is None:
            return out
        for child in scheme:
            name = child.tag.rsplit("}", 1)[-1]
            hexv = None
            srgb = child.find("a:srgbClr", NS)
            if srgb is not None:
                hexv = srgb.get("val")
            else:
                sysc = child.find("a:sysClr", NS)
                if sysc is not None:
                    hexv = sysc.get("lastClr")
            if hexv and len(hexv) == 6:
                try:
                    out[name] = (
                        int(hexv[0:2], 16),
                        int(hexv[2:4], 16),
                        int(hexv[4:6], 16),
                    )
                except ValueError:
                    pass
        aliases = {"bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2"}
        clr_map = master._element.find("p:clrMap", NS)
        if clr_map is not None:
            for name in aliases:
                if clr_map.get(name):
                    aliases[name] = clr_map.get(name)
        for alias, target in aliases.items():
            if target in out:
                out[alias] = out[target]
    except Exception:
        pass
    return out


def _bg_state(owner, scheme_colors):
    """<p:bg> を持つ持ち主（slide / layout / master）の背景の状態を返す。持たなければ None。

    単色の RGB が取れれば ("rgb", 色)。背景参照（bgRef）や単色のテーマ色参照は、
    テーマの色定義から解決できればその色。グラデーション・画像など色が
    決められないものは ("painted", None)。
    背景は XML から直接読む。python-pptx の background.fill に触ると、bgRef 形式の
    背景が bgPr 形式に作り替えられてテーマ参照が消えてしまうため。
    """
    csld = owner._element.find("p:cSld", NS)
    bg = csld.find("p:bg", NS) if csld is not None else None
    if bg is None:
        return None
    bgpr = bg.find("p:bgPr", NS)
    if bgpr is not None:
        rgb = _srgb_of_props(bgpr)
        if rgb is not None:
            return ("rgb", rgb)
    for path in ("p:bgRef/a:schemeClr", "p:bgPr/a:solidFill/a:schemeClr"):
        sc = bg.find(path, NS)
        if sc is not None and sc.get("val") in scheme_colors:
            return ("rgb", scheme_colors[sc.get("val")])
    return ("painted", None)


def slide_backgrounds(prs):
    """スライドごとの背景の状態 (種別, 色) のリスト。

    背景は、定義している最初の持ち主（スライド → レイアウト → マスター）で決まる。
    ("rgb", 色): 単色として色が決まった（テーマの色参照はテーマの定義から解決する）
    ("painted", None): 背景はあるが色が決められない（グラデーション・画像など）
    ("none", None): どこにも背景の定義が無い
    """
    out = []
    scheme_cache = {}
    for slide in prs.slides:
        master = slide.slide_layout.slide_master
        key = id(master.part)
        if key not in scheme_cache:
            scheme_cache[key] = _theme_scheme_colors(master)
        scheme = scheme_cache[key]
        state = ("none", None)
        for owner in (slide, slide.slide_layout, master):
            st = _bg_state(owner, scheme)
            if st is not None:
                state = st
                break
        out.append(state)
    return out


def _effective_background(elems, idx, unit, slide_bg_state, cfg):
    """文字ユニットの背景色 (r, g, b) を返す。色が決められなければ None（判定の対象外）。

    決める順序: ユニット自身の塗り（表のセルや図形自身の単色）→ 文字の下に
    ある図形の単色の塗り → スライドの背景 → 白（背景の定義が無い場合）。
    画像やグラデーションなど色が決められないものが文字の背景にある場合は
    None を返し、そのユニットは判定しない。
    """
    if unit.bg_state is not None:
        kind, color = unit.bg_state
        return color if kind == "rgb" else None
    e = elems[idx]
    area = e.w * e.h
    if area > 0:
        cover = cfg.get("contrast_cover_ratio", 0.8)
        for j in range(idx - 1, -1, -1):
            b = elems[j]
            ix = min(e.x + e.w, b.x + b.w) - max(e.x, b.x)
            iy = min(e.y + e.h, b.y + b.h) - max(e.y, b.y)
            if ix <= 0 or iy <= 0 or ix * iy < area * cover:
                continue
            kind, color = _fill_state(b)
            if kind == "rgb":
                return color
            if kind == "painted":
                return None
            # "none" の図形は塗りが無く透けるので、さらに下を探す
    kind, color = slide_bg_state
    if kind == "rgb":
        return color
    if kind == "painted":
        return None
    return (255, 255, 255)


def judge(slides, cfg, slide_w, slide_h, backgrounds=None):
    findings = {i: [] for i in range(1, len(slides) + 1)}
    stats = {
        "chars_total": 0,
        "chars_unresolved": 0,
        "unresolved": [],
        "by_src": {},
        "notes": [],
        "labels": [],
        "units": 0,
    }
    floors = {"body": cfg["body_min_pt"], "label": cfg.get("label_min_pt", 18)}
    body_pairs_all = []  # 資料全体の本文の (サイズ, 文字数)。全体の中央値用

    for sno, elems in enumerate(slides, 1):
        body_units = []  # このスライドの本文ユニット (要素, ユニット, 最小サイズ, 根拠)
        body_pairs = []  # このスライドの本文の (サイズ, 文字数)。中央値用
        label_median = slide_body_median(elems, cfg)
        bg_state = backgrounds[sno - 1] if backgrounds else ("none", None)
        for eidx, e in enumerate(elems):
            for u in e.units:
                if not u.text.strip():
                    continue
                stats["units"] += 1
                role, why = classify(e, u, cfg, label_median)
                floor_u = floors.get(role)
                eff_bg = _effective_background(elems, eidx, u, bg_state, cfg)
                small_chars = total = 0
                sizes = []
                contrast_worst = (
                    None  # (比, 文字色, しきい値)。しきい値未満で最も低いもの
                )
                for p in u.paragraphs:
                    for (n, size, src), color in zip(p["runs"], p["colors"]):
                        total += n
                        stats["chars_total"] += n
                        if eff_bg is not None and color is not None:
                            eff_c = size * u.scale if size is not None else None
                            limit_c = (
                                cfg.get("contrast_min_ratio_large", 3.0)
                                if eff_c is not None
                                and eff_c >= cfg.get("contrast_large_pt", 18)
                                else cfg.get("contrast_min_ratio", 4.5)
                            )
                            ratio_c = _contrast_ratio(color, eff_bg)
                            if ratio_c < limit_c and (
                                contrast_worst is None or ratio_c < contrast_worst[0]
                            ):
                                contrast_worst = (ratio_c, color, limit_c)
                        if size is None:
                            stats["chars_unresolved"] += n
                            stats["unresolved"].append(
                                (sno, e.name, u.label, p["text"][:20])
                            )
                            continue
                        eff = size * u.scale
                        stats["by_src"][src] = stats["by_src"].get(src, 0) + n
                        sizes.append(eff)
                        if role == "body":
                            body_pairs.append((eff, n))
                        if floor_u is not None and eff < floor_u:
                            small_chars += n
                if role == "body" and sizes:
                    body_units.append((e, u, min(sizes), why))
                if role == "note":
                    stats["notes"].append(
                        (
                            sno,
                            e.name,
                            u.label,
                            why,
                            min(sizes) if sizes else None,
                            u.text[:20],
                        )
                    )
                if role == "label":
                    stats["labels"].append(
                        (
                            sno,
                            e.name,
                            u.label,
                            why,
                            min(sizes) if sizes else None,
                            u.text[:20],
                        )
                    )
                if floor_u is not None and small_chars:
                    scale_note = (
                        "（自動縮小 %d%% 適用後の値）" % round(u.scale * 100)
                        if u.scale != 1.0
                        else ""
                    )
                    findings[sno].append(
                        {
                            "type": "小さい文字",
                            "elem": e,
                            "label": u.label,
                            "head": u.text[:24].replace("\n", " "),
                            "detail": "最小 %.1fpt（下限 %gpt、%d/%d文字が下限未満）%s。役割=%s（%s）"
                            % (
                                min(sizes),
                                floor_u,
                                small_chars,
                                total,
                                scale_note,
                                ROLE_JP[role],
                                why,
                            ),
                        }
                    )
                if contrast_worst is not None:
                    ratio_c, color_c, limit_c = contrast_worst
                    findings[sno].append(
                        {
                            "type": "コントラストが低い",
                            "elem": e,
                            "label": u.label,
                            "head": u.text[:24].replace("\n", " "),
                            "detail": "文字色 #%02X%02X%02X / 背景色 #%02X%02X%02X / コントラスト比 %.2f（しきい値 %.1f 未満）"
                            % (
                                color_c[0],
                                color_c[1],
                                color_c[2],
                                eff_bg[0],
                                eff_bg[1],
                                eff_bg[2],
                                ratio_c,
                                limit_c,
                            ),
                        }
                    )

        # 周囲とのバランス（相対判定）: そのスライドの本文の中央値と比べて
        # 小さい文字を出す。絶対値の下限で既に指摘した文字と重ねないため、
        # 対象は下限を満たしている本文だけにする
        body_pairs_all.extend(body_pairs)
        if len(body_units) >= cfg.get("relative_min_body_units", 3):
            med = _weighted_median(body_pairs)
            ratio = cfg.get("relative_min_ratio", 0.7)
            if med is not None:
                for e, u, unit_min, why in body_units:
                    if floors["body"] <= unit_min < med * ratio:
                        findings[sno].append(
                            {
                                "type": "周囲より小さい",
                                "elem": e,
                                "label": u.label,
                                "head": u.text[:24].replace("\n", " "),
                                "detail": "最小 %.1fpt はスライドの本文の中央値 %.1fpt の %.0f%% で、しきい値 %.0f%% を下回る。本文の下限 %gpt は満たしている。役割=本文（%s）"
                                % (
                                    unit_min,
                                    med,
                                    unit_min / med * 100,
                                    ratio * 100,
                                    floors["body"],
                                    why,
                                ),
                            }
                        )

        # はみ出し（許容は要素種別で分ける: 画像は広く、それ以外は従来通り）
        for e in elems:
            if e.kind == "線・コネクタ":
                continue
            tol = (
                cfg.get("image_out_of_bounds_tolerance_mm", 2.0)
                if e.kind == "図(画像)"
                else cfg["out_of_bounds_tolerance_mm"]
            ) * EMU_PER_MM
            over = []
            if e.x < -tol:
                over.append("左 %.1fmm" % mm(-e.x))
            if e.y < -tol:
                over.append("上 %.1fmm" % mm(-e.y))
            if e.x + e.w > slide_w + tol:
                over.append("右 %.1fmm" % mm(e.x + e.w - slide_w))
            if e.y + e.h > slide_h + tol:
                over.append("下 %.1fmm" % mm(e.y + e.h - slide_h))
            if over:
                findings[sno].append(
                    {
                        "type": "はみ出し",
                        "elem": e,
                        "label": "",
                        "head": (e.text or e.name)[:24].replace("\n", " "),
                        "detail": "スライド範囲 %.0fx%.0fmm に対し、位置(%.1f, %.1f) 大きさ %.1fx%.1fmm。はみ出し: %s"
                        % (
                            mm(slide_w),
                            mm(slide_h),
                            mm(e.x),
                            mm(e.y),
                            mm(e.w),
                            mm(e.h),
                            "、".join(over),
                        ),
                    }
                )

        # 重なり
        cand = [
            e
            for e in elems
            if e.kind != "線・コネクタ"
            and (e.has_text or e.kind in ("図(画像)", "表", "グラフ", "SmartArt"))
        ]
        for i in range(len(cand)):
            for j in range(i + 1, len(cand)):
                a, b = cand[i], cand[j]
                if a.kind == "図(画像)" and b.kind == "図(画像)":
                    continue  # 画像同士の重なりはレイヤーのデザインとして対象外
                ix = min(a.x + a.w, b.x + b.w) - max(a.x, b.x)
                iy = min(a.y + a.h, b.y + b.h) - max(a.y, b.y)
                if ix <= 0 or iy <= 0:
                    continue
                if mm(ix) < cfg["overlap_min_mm"] or mm(iy) < cfg["overlap_min_mm"]:
                    continue
                small = min(a.w * a.h, b.w * b.h) or 1
                ratio = ix * iy / small
                if ratio < cfg["overlap_min_ratio"]:
                    continue
                findings[sno].append(
                    {
                        "type": "重なり",
                        "elem": a,
                        "label": "",
                        "head": (a.text or a.name)[:24].replace("\n", " "),
                        "detail": "「%s」(%s) と「%s」(%s) が重なる。重なり %.1fx%.1fmm、小さい方の面積の %.0f%%"
                        % (
                            (a.text or a.name)[:16].replace("\n", " "),
                            a.kind,
                            (b.text or b.name)[:16].replace("\n", " "),
                            b.kind,
                            mm(ix),
                            mm(iy),
                            ratio * 100,
                        ),
                    }
                )
    stats["body_median_all"] = _weighted_median(body_pairs_all)
    return findings, stats


# ---------------------------------------------------------------- 構成の指摘（指摘のみ）
def judge_structure(slides, cfg, slide_w, slide_h):
    """構成の指摘を返す。通常の自動調整（--fix）では直さない指摘。

    judge() の findings と同じ形式で返すが、既存の「指摘の合計」には含めない
    別枠にする（資料全体の中央値の通知と同じ扱い）。文章の意味を読む判断は
    扱わず、形のパターンで見分けられるものだけを見る。直す場合は、レポートの
    番号で選んで --structure-apply（または GUI）から適用する。
    """
    out = {i: [] for i in range(1, len(slides) + 1)}
    for sno, elems in enumerate(slides, 1):
        body_median = slide_body_median(elems, cfg)
        out[sno].extend(_structure_repeats(elems, cfg, body_median))
        out[sno].extend(_structure_orphans(elems, cfg, body_median))
        out[sno].extend(_structure_color_mismatch(elems, cfg))
        out[sno].extend(_structure_footnotes(elems, cfg, slide_h))
    return out


def _structure_numbered(structure_findings):
    """構成の指摘に、資料全体を通した連番（1から）を付けた並びを返す。

    並び順はスライド順、その中では judge_structure の検出順に固定する。
    レポートの表示と --structure-apply / GUI の指定で、同じ番号が同じ指摘を
    指すための物差し。
    """
    out = []
    no = 0
    for sno in sorted(structure_findings):
        for f in structure_findings[sno]:
            no += 1
            out.append((no, sno, f))
    return out


def _norm_short_text(text):
    """繰り返しの比較用に、前後の空白を除き、連続する空白を1つに圧縮した形にする。"""
    return re.sub(r"\s+", " ", text.strip())


def _structure_repeats(elems, cfg, body_median):
    """同じ短い文が別々の要素に繰り返し出ている箇所を返す。

    「リスク」のような小見出しが、並んだ要素ごとに繰り返される形を拾う。
    数えるのは段落単位で、同じ要素の中の重複は1回と数える。表の中と、
    注記に分類されたユニットは対象にしない。
    """
    min_count = cfg.get("structure_repeat_min", 3)
    places = {}  # 正規化した文 -> {要素の位置: (要素, ユニット)}
    order = []
    for eidx, e in enumerate(elems):
        if e.kind == "表":
            continue
        for u in e.units:
            if not u.text.strip():
                continue
            role, _why = classify(e, u, cfg, body_median)
            if role == "note":
                continue
            for p in u.paragraphs:
                t = _norm_short_text(p["text"])
                if not 2 <= len(t) <= 30:
                    continue
                if re.match(cfg["page_number_regex"], t):
                    continue
                if not any(c.isalnum() and not c.isdigit() for c in t):
                    continue  # 数字・記号だけのものは見出しとして扱わない
                if t not in places:
                    places[t] = {}
                    order.append(t)
                places[t].setdefault(eidx, (e, u))
    out = []
    for t in order:
        found = places[t]
        if len(found) < min_count:
            continue
        first_e, first_u = next(iter(found.values()))
        names = "、".join(e.name for e, _u in found.values())
        out.append(
            {
                "type": "見出しの繰り返し",
                "elem": first_e,
                "label": first_u.label,
                "head": t,
                "detail": "「%s」という短い文が、このスライドの %d 個の要素に繰り返し出ている（%s）。共通の見出しとしてまとめられないか"
                % (t, len(found), names),
            }
        )
    return out


def _paragraph_segments(unit, para):
    """段落を明示改行で区切った (文字, 実効サイズpt) の並びのリストを返す。

    サイズが解決できない run が段落にあれば None（推定できないため）。
    明示改行（a:br）は段落の text の中に改行として入っている。run の文字列を
    つなげたものは text から改行を除いたものと一致するので、改行の位置は
    text 側の区切りの長さで復元する。
    """
    stream = []
    for (_n, size, _src), (_el, txt) in zip(para["runs"], para["run_els"]):
        if size is None:
            return None
        eff = size * unit.scale
        stream.extend((ch, eff) for ch in txt)
    seg_lens = [len(s) for s in para["text"].split("\n")]
    if sum(seg_lens) != len(stream):
        seg_lens = [len(stream)]
    segs = []
    pos = 0
    for ln in seg_lens:
        segs.append(stream[pos : pos + ln])
        pos += ln
    return segs


def _wrapped_line_counts(seg, inner_w_pt):
    """(文字, サイズpt) の並びを inner_w_pt の幅で折り返したときの、行ごとの文字数を返す。"""
    counts = []
    cur_n = 0
    cur_w = 0.0
    for ch, size in seg:
        w = _char_factor(ch) * size
        if cur_n and cur_w + w > inner_w_pt:
            counts.append(cur_n)
            cur_n = 0
            cur_w = 0.0
        cur_n += 1
        cur_w += w
    counts.append(cur_n)
    return counts


def _structure_orphans(elems, cfg, body_median):
    """折り返しの最後の行に数文字だけ残る段落を返す。

    折り返しは _est_height_mm と同じ近似（文字幅の係数と FIT_SAFETY）で
    推定する。著者が明示した改行より手前は著者の判断として見ず、
    最後の区切りが折り返した結果だけを見る。対象は本文とラベル。
    """
    max_chars = cfg.get("structure_orphan_max_chars", 3)
    min_chars = cfg.get("structure_orphan_min_chars", 12)
    out = []
    for e in elems:
        for u in e.units:
            if not u.text.strip() or u.inner_w <= 0:
                continue
            role, _why = classify(e, u, cfg, body_median)
            if role not in ("body", "label"):
                continue
            inner_w_pt = mm(u.inner_w) * FIT_SAFETY / PT_TO_MM
            for p in u.paragraphs:
                segs = _paragraph_segments(u, p)
                if segs is None:
                    continue
                total = sum(len(s) for s in segs)
                if total < min_chars or not segs[-1]:
                    continue
                counts = _wrapped_line_counts(segs[-1], inner_w_pt)
                if len(counts) < 2 or counts[-1] > max_chars:
                    continue
                seg_text = p["text"].split("\n")[-1]
                orphan = seg_text[len(seg_text) - counts[-1] :]
                out.append(
                    {
                        "type": "改行の孤立",
                        "elem": e,
                        "label": u.label,
                        "head": u.text[:24].replace("\n", " "),
                        "detail": "段落（全 %d 文字）の折り返しを推定すると、最後の行が「%s」の %d 文字だけになる"
                        % (total, orphan, counts[-1]),
                    }
                )
                break  # ユニットごとに最初の該当段落で1件
    return out


def _color_groups(elems, cfg):
    """横か縦に並んだ同形の要素の組を (軸, [(要素, 色)]) で順に返す。

    色の不揃いの判定と、その修正で並びの組を組み直すときの共通の物差し。
    全員の単色塗りが RGB で決まる要素だけを候補にし、大きさと並び位置が
    近いものを組にする。重なっている要素同士は「並んだ」とはみなさない。
    """
    min_n = cfg.get("structure_sibling_min", 3)
    tol = 5.0 * EMU_PER_MM
    cand = []
    for e in elems:
        if not e.has_text:
            continue
        kind, color = _fill_state(e)
        if kind == "rgb":
            cand.append((e, color))
    seen_groups = set()
    for axis in ("row", "col"):
        for seed, _c in cand:
            if axis == "row":
                near = [(e, c) for e, c in cand if abs(e.y - seed.y) <= tol]
            else:
                near = [(e, c) for e, c in cand if abs(e.x - seed.x) <= tol]
            if len(near) < min_n:
                continue
            ws = sorted(e.w for e, _c in near)
            hs = sorted(e.h for e, _c in near)
            ps = sorted((e.y if axis == "row" else e.x) for e, _c in near)
            med_w, med_h = ws[len(ws) // 2], hs[len(hs) // 2]
            med_p = ps[len(ps) // 2]
            group = [
                (e, c)
                for e, c in near
                if abs(e.w - med_w) <= med_w * 0.25
                and abs(e.h - med_h) <= med_h * 0.25
                and abs((e.y if axis == "row" else e.x) - med_p) <= tol
            ]
            if len(group) < min_n:
                continue
            key = frozenset(id(e) for e, _c in group)
            if key in seen_groups:
                continue
            seen_groups.add(key)
            overlapped = False
            for i in range(len(group)):
                for j in range(i + 1, len(group)):
                    a, b = group[i][0], group[j][0]
                    ix = min(a.x + a.w, b.x + b.w) - max(a.x, b.x)
                    iy = min(a.y + a.h, b.y + b.h) - max(a.y, b.y)
                    if ix > EMU_PER_MM and iy > EMU_PER_MM:
                        overlapped = True
            if overlapped:
                continue
            yield axis, group


def _structure_color_mismatch(elems, cfg):
    """横か縦に並んだ同形の要素のうち、1つだけ塗りの色が違うものを返す。

    全員の単色塗りが RGB で決まる場合だけを見る。全員が同じ色の場合と
    全員ばらばらの場合は、揃えたのか意図して変えたのかを決められないので
    対象にしない。
    """
    out = []
    flagged = set()
    for axis, group in _color_groups(elems, cfg):
        counts = {}
        for _e, c in group:
            counts[c] = counts.get(c, 0) + 1
        if len(counts) != 2:
            continue
        odd_color = next((c for c, n in counts.items() if n == 1), None)
        if odd_color is None:
            continue
        main_color = next(c for c, n in counts.items() if n == len(group) - 1)
        odd_e = next(e for e, c in group if c == odd_color)
        if id(odd_e) in flagged:
            continue
        flagged.add(id(odd_e))
        unit = next(
            (u for u in odd_e.units if u.text.strip()),
            odd_e.units[0] if odd_e.units else None,
        )
        out.append(
            {
                "type": "色の不揃い",
                "elem": odd_e,
                "label": unit.label if unit else "",
                "head": odd_e.text[:24].replace("\n", " "),
                "detail": "%sに並んだ %d 個の要素のうち、この要素だけ塗りの色が違う（この要素 #%02X%02X%02X / 他の %d 個は #%02X%02X%02X）"
                % (
                    "横" if axis == "row" else "縦",
                    len(group),
                    odd_color[0],
                    odd_color[1],
                    odd_color[2],
                    len(group) - 1,
                    main_color[0],
                    main_color[1],
                    main_color[2],
                ),
            }
        )
    return out


def _structure_footnotes(elems, cfg, slide_h):
    """スライドの下端にある、長いのに小さい文字の文を返す。役割は問わない。"""
    ratio = cfg.get("structure_footnote_top_ratio", 0.85)
    min_chars = cfg.get("structure_footnote_min_chars", 80)
    floor = cfg["body_min_pt"]
    out = []
    for e in elems:
        if e.y <= slide_h * ratio:
            continue
        for u in e.units:
            if not u.text.strip():
                continue
            size = _min_size(u)
            if size is None or size >= floor:
                continue
            n = sum(n for p in u.paragraphs for (n, _s, _src) in p["runs"])
            if n < min_chars:
                continue
            out.append(
                {
                    "type": "長い注釈",
                    "elem": e,
                    "label": u.label,
                    "head": u.text[:24].replace("\n", " "),
                    "detail": "スライドの下端に、最小 %.1fpt（本文の下限 %gpt 未満）で %d 文字の文がある。この長さの文は、小さい文字のままでは読みにくい"
                    % (size, floor, n),
                }
            )
    return out


# ---------------------------------------------------------------- 自動調整（フェーズ2）
PT_TO_MM = 0.3528
LINE_HEIGHT = 1.2  # 行高はフォントサイズの1.2倍として推定
FIT_SAFETY = 0.98  # 推定を安全側に倒すため、使える内寸を少し小さく見る


def _char_factor(ch):
    """文字幅の近似係数（フォントサイズ比）。全角=1.0、半角=0.55。"""
    return 1.0 if unicodedata.east_asian_width(ch) in ("W", "F") else 0.55


def _est_height_mm(unit, raise_to):
    """raise_to(pt) 未満の run を raise_to まで上げた場合のテキスト推定高さ(mm)。

    段落ごとに、文字幅の合計から折り返し行数を推定し、
    行数 × 最大サイズ × 行高 で高さを積む。改行（a:br）の分は最低行数として数える。
    """
    inner_w = mm(unit.inner_w) * FIT_SAFETY
    total = 0.0
    for p in unit.paragraphs:
        if not p["runs"]:
            total += raise_to * LINE_HEIGHT * PT_TO_MM
            continue
        width_pt = 0.0
        max_size = 0.0
        for (n, size, src), (el, txt) in zip(p["runs"], p["run_els"]):
            if size is None:
                continue
            new = max(size * unit.scale, raise_to)
            max_size = max(max_size, new)
            width_pt += sum(_char_factor(c) for c in txt) * new
        if max_size == 0:
            total += raise_to * LINE_HEIGHT * PT_TO_MM
            continue
        wrapped = math.ceil(width_pt * PT_TO_MM / inner_w) if inner_w > 0 else 1
        lines = max(p["text"].count("\n") + 1, wrapped)
        total += lines * max_size * LINE_HEIGHT * PT_TO_MM
    return total


def _fit_size(unit, floor):
    """floor 以下で枠に収まる最大のサイズ(pt)。現状の最小サイズ以下なら None。"""
    cur_min = _min_size(unit)
    if cur_min is None or cur_min >= floor:
        return None
    inner_h = mm(unit.inner_h) * FIT_SAFETY
    cand = float(floor)
    while cand > cur_min:
        if _est_height_mm(unit, cand) <= inner_h:
            return cand
        cand -= 0.5
    return None


def _apply_size(unit, target_eff):
    """target_eff(pt) 未満の run に、run 単位の明示サイズを設定する。"""
    for p in unit.paragraphs:
        for (n, size, src), (el, txt) in zip(p["runs"], p["run_els"]):
            if size is None:
                continue
            if size * unit.scale < target_eff - 1e-9:
                raw = target_eff / unit.scale if unit.scale else target_eff
                rpr = el.find("a:rPr", NS)
                if rpr is None:
                    rpr = el.makeelement(A + "rPr", {})
                    el.insert(0, rpr)
                rpr.set("sz", str(int(round(raw * 100))))


def _needed_heading_size(elems, heading, heading_size, cap, cfg):
    """本文の拡大目標が上限を超える場合に、見出しに必要な大きさ(pt)を返す。

    対象が無ければ None。本文ごとに「拡大目標 + マージン」と「元の比率を
    保つ大きさ（拡大目標 × 見出しサイズ ÷ 本文の現在の最小サイズ）」の
    大きい方を必要値とし、スライド内ではその最大値を採る。拡大目標は
    下限と枠に収まる範囲で決まる大きさ（_fit_size の結果）で、枠の都合で
    そもそも拡大できない本文は対象にしない。
    """
    margin = cfg.get("heading_margin_pt", 2)
    floor = cfg["body_min_pt"]
    median = slide_body_median(elems, cfg)
    needed = None
    for e in elems:
        if e is heading:
            continue
        for u in e.units:
            if not u.text.strip():
                continue
            role, _why = classify(e, u, cfg, median)
            if role != "body":
                continue
            cur_min = _min_size(u)
            if cur_min is None or cur_min >= floor:
                continue
            target = _fit_size(u, floor)
            if target is None or target <= cap:
                continue
            cand = max(target + margin, target * heading_size / cur_min)
            needed = cand if needed is None else max(needed, cand)
    return needed


def _raise_heading(heading, cur_size, needed):
    """見出しを needed(pt) まで、枠に収まる範囲で上げる。上げた後の大きさを返す。

    収まりの推定は本文の拡大と同じ _est_height_mm を使い、needed から
    0.5pt 刻みで下げて、現在の大きさより大きく収まる最大の大きさを探す。
    現在の大きさ以下しか収まらない場合は上げず、None を返す。
    """
    units = [u for u in heading.units if u.text.strip()]
    if not units:
        return None
    cand = float(needed)
    while cand > cur_size:
        if all(_est_height_mm(u, cand) <= mm(u.inner_h) * FIT_SAFETY for u in units):
            for u in units:
                _apply_size(u, cand)
            return cand
        cand -= 0.5
    return None


def _move(e, dx, dy):
    """要素を絶対座標で (dx, dy) EMU 動かす。グループ内はローカル座標へ換算する。"""
    if dx:
        e.shape.left = int(round(e.shape.left + dx / e.tf.sx))
        e.x += dx
    if dy:
        e.shape.top = int(round(e.shape.top + dy / e.tf.sy))
        e.y += dy


def _tol_for(e, cfg):
    """はみ出しの許容(EMU)。画像は広く、それ以外は従来通り。"""
    return (
        cfg.get("image_out_of_bounds_tolerance_mm", 2.0)
        if e.kind == "図(画像)"
        else cfg["out_of_bounds_tolerance_mm"]
    ) * EMU_PER_MM


def apply_fixes(slides, cfg, slide_w, slide_h):
    """サイズ拡大・はみ出し解消・重なり解消を行い、変更の一覧を返す。"""
    changes = []

    def rec(sno, type_, e, label, detail, head=None):
        changes.append(
            {
                "slide": sno,
                "type": type_,
                "elem": e.name,
                "label": label,
                "head": (head if head is not None else (e.text or e.name))[:24].replace(
                    "\n", " "
                ),
                "detail": detail,
            }
        )

    floors = {"body": cfg["body_min_pt"], "label": cfg.get("label_min_pt", 18)}
    margin = cfg.get("heading_margin_pt", 2)
    for sno, elems in enumerate(slides, 1):
        label_median = slide_body_median(elems, cfg)
        # 階層の保護: 見出しのサイズからマージンを引いた値を、本文を拡大する
        # ときの上限にする。見出しの決め方は分割・並び替えと同じ _split_heading
        # に揃える
        heading = _split_heading(elems, cfg)
        heading_size = _max_size(heading) if heading is not None else None
        cap = heading_size - margin if heading_size is not None else None

        # 見出し側を上げる調整: 本文の拡大目標が上限を超えるスライドでは、
        # 本文を上限で止める前に、見出しを必要な大きさまで上げて上限自体を
        # 引き上げる。見出しが上がった場合は、上限は上がった後の大きさで
        # 計算し直し、従来の階層の保護をそのまま適用する
        if heading is not None and heading_size is not None:
            needed = _needed_heading_size(elems, heading, heading_size, cap, cfg)
            if needed is not None:
                raised = _raise_heading(heading, heading_size, needed)
                if raised is not None:
                    detail = (
                        "見出し %.1fpt → %.1fpt（本文を下限 %gpt まで上げるため。"
                        "見出しと本文の元の比率を保つ大きさ）"
                        % (heading_size, raised, floors["body"])
                    )
                    if raised < needed:
                        detail += (
                            "。必要だった %.1fpt までは枠に収まらず、収まる最大にした"
                            % needed
                        )
                    label = next((u.label for u in heading.units if u.text.strip()), "")
                    rec(sno, "サイズ", heading, label, detail)
                    heading_size = raised
                    cap = heading_size - margin

        # 1. サイズの拡大（note と title は触らない）
        for e in elems:
            for u in e.units:
                if not u.text.strip():
                    continue
                role, why = classify(e, u, cfg, label_median)
                floor = floors.get(role)
                if floor is None:
                    continue
                cur_min = _min_size(u)
                if cur_min is None:
                    continue
                if cur_min >= floor:
                    continue
                desired = floor
                if role == "body" and cap is not None and e is not heading:
                    desired = min(floor, cap)
                    if desired <= cur_min:
                        rec(
                            sno,
                            "未解消",
                            e,
                            u.label,
                            "拡大の上限 %.1fpt（見出し %.1fpt − マージン %gpt）が現在の最小 %.1fpt 以下になるため、拡大しなかった。役割=本文（%s）"
                            % (cap, heading_size, margin, cur_min, why),
                            head=u.text,
                        )
                        continue
                target = _fit_size(u, desired)
                if target is None:
                    continue
                _apply_size(u, target)
                detail = (
                    "最小 %.1fpt → %.1fpt（役割=%s、下限 %gpt。枠に収まる範囲で拡大）"
                    % (cur_min, target, ROLE_JP[role], floor)
                )
                if desired < floor:
                    detail += (
                        "。見出し %.1fpt より大きくしないため、上限 %.1fpt に抑えた"
                        % (
                            heading_size,
                            cap,
                        )
                    )
                rec(sno, "サイズ", e, u.label, detail, head=u.text)

        # 2. はみ出しの解消（スライド範囲内へ平行移動で戻す）
        for e in elems:
            if e.kind == "線・コネクタ" or e.shape is None:
                continue
            tol = _tol_for(e, cfg)
            dx = dy = 0.0
            if e.x < -tol:
                dx = -e.x
            elif e.x + e.w > slide_w + tol:
                dx = slide_w - (e.x + e.w)
            if e.y < -tol:
                dy = -e.y
            elif e.y + e.h > slide_h + tol:
                dy = slide_h - (e.y + e.h)
            if dx or dy:
                parts = []
                if dx < 0:
                    parts.append("左へ %.1fmm" % mm(-dx))
                if dx > 0:
                    parts.append("右へ %.1fmm" % mm(dx))
                if dy < 0:
                    parts.append("上へ %.1fmm" % mm(-dy))
                if dy > 0:
                    parts.append("下へ %.1fmm" % mm(dy))
                _move(e, dx, dy)
                rec(
                    sno,
                    "移動（はみ出し解消）",
                    e,
                    "",
                    "スライド範囲内へ戻すため %s 移動" % "、".join(parts),
                )

        # 3. 重なりの解消（テキストを持つ要素同士。後ろの要素を下→右の順に試す）
        cand = [e for e in elems if e.has_text and e.shape is not None]
        for i in range(len(cand)):
            for j in range(i + 1, len(cand)):
                a, b = cand[i], cand[j]
                ix = min(a.x + a.w, b.x + b.w) - max(a.x, b.x)
                iy = min(a.y + a.h, b.y + b.h) - max(a.y, b.y)
                if ix <= 0 or iy <= 0:
                    continue
                if mm(ix) < cfg["overlap_min_mm"] or mm(iy) < cfg["overlap_min_mm"]:
                    continue
                small = min(a.w * a.h, b.w * b.h) or 1
                if ix * iy / small < cfg["overlap_min_ratio"]:
                    continue
                pair = "「%s」と「%s」" % (
                    (a.text or a.name)[:16].replace("\n", " "),
                    (b.text or b.name)[:16].replace("\n", " "),
                )
                dy = (a.y + a.h) - b.y
                if b.y + dy + b.h <= slide_h:
                    _move(b, 0.0, dy)
                    rec(
                        sno,
                        "移動（重なり解消）",
                        b,
                        "",
                        "%s の重なりを解消するため 下へ %.1fmm 移動" % (pair, mm(dy)),
                    )
                    continue
                dx = (a.x + a.w) - b.x
                if b.x + dx + b.w <= slide_w:
                    _move(b, dx, 0.0)
                    rec(
                        sno,
                        "移動（重なり解消）",
                        b,
                        "",
                        "%s の重なりを解消するため 右へ %.1fmm 移動" % (pair, mm(dx)),
                    )
                    continue
                rec(
                    sno,
                    "未解消",
                    b,
                    "",
                    "%s の重なりを、スライド内に収まる範囲では解消できなかった" % pair,
                )
    return changes


STRUCTURE_STATUS_JP = {
    "適用": "修正した",
    "解消済み": "すでに解消していたため対象外",
    "未解消": "解消できなかった",
    "対象外": "対象外",
}


def build_changes_report(path, out_name, changes, structure_results=None):
    L = []
    L.append("# ReSlide 変更一覧: %s" % Path(path).name)
    L.append("")
    L.append("- 出力ファイル: %s" % out_name)
    n_size = sum(1 for c in changes if c["type"] == "サイズ")
    n_move = sum(1 for c in changes if c["type"].startswith("移動"))
    n_fail = sum(1 for c in changes if c["type"] == "未解消")
    n_split = sum(1 for c in changes if c["type"] == "分割")
    L.append(
        "- 変更の合計: %d件（サイズ %d件 / 移動 %d件 / 解消できなかった %d件）"
        % (len(changes), n_size, n_move, n_fail)
    )
    if n_split:
        L.append("- うちページ分割: %d件" % n_split)
    n_reorder = sum(1 for c in changes if c["type"] == "並び替え")
    if n_reorder:
        L.append("- うち並び替え: %d件" % n_reorder)
    if structure_results:
        n_ok = sum(1 for r in structure_results if r["status"] == "適用")
        L.append(
            "- うち構成の指摘の修正: %d件を指定（修正した %d件）"
            % (len(structure_results), n_ok)
        )
    if not changes and not structure_results:
        L.append("")
        L.append("変更なし")
        return "\n".join(L) + "\n"
    by_slide = {}
    for c in changes:
        by_slide.setdefault(c["slide"], []).append(c)
    for sno in sorted(by_slide):
        L.append("")
        L.append("### スライド %d（%d件）" % (sno, len(by_slide[sno])))
        for c in by_slide[sno]:
            L.append(
                "- [%s] %s%s 「%s」"
                % (
                    c["type"],
                    c["elem"],
                    (" " + c["label"]) if c["label"] else "",
                    c["head"],
                )
            )
            L.append("  - %s" % c["detail"])
    if structure_results:
        n_ok = sum(1 for r in structure_results if r["status"] == "適用")
        n_gone = sum(1 for r in structure_results if r["status"] == "解消済み")
        n_fail = sum(1 for r in structure_results if r["status"] == "未解消")
        n_na = sum(1 for r in structure_results if r["status"] == "対象外")
        L.append("")
        L.append("## 構成の指摘の修正（--structure-apply）")
        L.append("")
        L.append(
            "- 指定 %d件: 修正した %d件 / すでに解消していた %d件 / 解消できなかった %d件 / 対象外 %d件"
            % (len(structure_results), n_ok, n_gone, n_fail, n_na)
        )
        for r in structure_results:
            L.append(
                "- 番号%d [%s] スライド%d %s%s 「%s」"
                % (
                    r["no"],
                    r["type"],
                    r["slide"],
                    r["elem"],
                    (" " + r["label"]) if r["label"] else "",
                    r["head"],
                )
            )
            L.append("  - 結果: %s" % STRUCTURE_STATUS_JP[r["status"]])
            L.append("  - %s" % r["detail"])
    return "\n".join(L) + "\n"


# ---------------------------------------------------------------- ページ分割（フェーズ3）
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
R_ATTR_PREFIX = "{%s}" % R_NS  # r:embed / r:id など、リレーションシップ参照の属性

# 分割で要素ごと移せる種類（内部は分割しない）
SPLIT_MOVABLE_KINDS = ("表", "グラフ", "SmartArt")


def _text_overlap_pairs(elems, cfg):
    """テキストを持つ要素同士の重なりペア数。

    物差しは判定（judge）の重なりと同じ設定（overlap_min_mm / overlap_min_ratio）。
    判定側の指摘には図との組み合わせも含まれるため、ここでは数え直す。
    """
    cand = [e for e in elems if e.has_text]
    pairs = 0
    for i in range(len(cand)):
        for j in range(i + 1, len(cand)):
            a, b = cand[i], cand[j]
            ix = min(a.x + a.w, b.x + b.w) - max(a.x, b.x)
            iy = min(a.y + a.h, b.y + b.h) - max(a.y, b.y)
            if ix <= 0 or iy <= 0:
                continue
            if mm(ix) < cfg["overlap_min_mm"] or mm(iy) < cfg["overlap_min_mm"]:
                continue
            small = min(a.w * a.h, b.w * b.h) or 1
            if ix * iy / small >= cfg["overlap_min_ratio"]:
                pairs += 1
    return pairs


def judge_split(slides, cfg):
    """スライドごとの詰め込み判定。実測値と候補かどうかを返す。

    候補になるのは、テキスト要素数・総文字数・テキスト同士の重なりの
    3つすべてが設定のしきい値以上のスライドだけ。要素数だけ多い図の
    ラベルだらけのスライドや、タイトルだけのスライドは対象にならない。
    """
    verdicts = []
    for sno, elems in enumerate(slides, 1):
        n_text = sum(1 for e in elems if e.has_text)
        chars = sum(len(u.text) for e in elems for u in e.units)
        pairs = _text_overlap_pairs(elems, cfg)
        verdicts.append(
            {
                "slide": sno,
                "text_elements": n_text,
                "chars": chars,
                "overlaps": pairs,
                "candidate": n_text >= cfg["split_min_text_elements"]
                and chars >= cfg["split_min_chars"]
                and pairs >= cfg["split_min_overlaps"],
            }
        )
    return verdicts


def _split_reason(v, cfg):
    """詰め込み判定の根拠（実測値としきい値の照合）を1行で返す。"""
    return (
        "テキスト要素 %d（しきい値 %d以上）/ 総文字数 %d（しきい値 %d以上）"
        "/ テキスト同士の重なり %d件（しきい値 %d件以上）→ %s"
        % (
            v["text_elements"],
            cfg["split_min_text_elements"],
            v["chars"],
            cfg["split_min_chars"],
            v["overlaps"],
            cfg["split_min_overlaps"],
            "分割候補" if v["candidate"] else "対象外",
        )
    )


def _max_size(e):
    """要素内の解決済みフォントサイズの最大値（実効値）。無ければ None。"""
    sizes = [
        size * u.scale
        for u in e.units
        for p in u.paragraphs
        for (n, size, src) in p["runs"]
        if size is not None
    ]
    return max(sizes) if sizes else None


def _split_heading(elems, cfg):
    """見出しに当たる要素。タイトル用プレースホルダを優先し、
    無ければ解決済みサイズが最大のテキスト要素を当てる。"""
    for e in elems:
        if e.has_text and e.ph_type in cfg["title_placeholder_types"]:
            return e
    best = None
    for e in elems:
        if not e.has_text:
            continue
        s = _max_size(e)
        if s is not None and (best is None or s > _max_size(best)):
            best = e
    return best


def _is_note_only(e, cfg, body_median=None):
    """テキストの役割がすべて注記（出典・ページ番号など）の要素か。

    注記はスライドの付属物として元のスライドに残す
    （フェーズ2で注記のサイズを触らないのと同じ扱い）。
    """
    seen = False
    for u in e.units:
        if not u.text.strip():
            continue
        seen = True
        role, _why = classify(e, u, cfg, body_median)
        if role != "note":
            return False
    return seen


def _split_movable(elems, heading, cfg):
    """分割で後半へ移す要素。読み順（上から、左から）に並べて返す。

    対象はスライド直下の要素だけ。グループ内の要素は親のグループごとで
    なければ移せないため対象にしない。画像は元のスライドに残し、
    表・グラフ・SmartArt は内部を割らず要素ごと移す。
    """
    median = slide_body_median(elems, cfg)
    out = []
    for e in elems:
        if e is heading or e.shape is None or e.in_group:
            continue
        if not (e.has_text or e.kind in SPLIT_MOVABLE_KINDS):
            continue
        if _is_note_only(e, cfg, median):
            continue
        out.append(e)
    out.sort(key=lambda e: (e.y, e.x))
    return out


def _retarget_rels(src_part, dst_part, element):
    """移した要素のリレーションシップ参照を、移し先のスライド側へ張り替える。

    画像やグラフ部品は r:embed / r:id 等の属性で部品を指している。要素の
    XML だけを別のスライドへ移すと参照が元のスライド側に残るため、同じ
    参照先へのリレーションシップを移し先にも作り、属性値を差し替える。
    """
    for el in element.iter():
        for attr in list(el.attrib):
            if not attr.startswith(R_ATTR_PREFIX):
                continue
            rid = el.get(attr)
            if not rid or rid not in src_part.rels:
                continue
            rel = src_part.rels[rid]
            if rel.is_external:
                new_rid = dst_part.relate_to(
                    rel.target_ref, rel.reltype, is_external=True
                )
            else:
                new_rid = dst_part.relate_to(rel.target_part, rel.reltype)
            el.set(attr, new_rid)


def _insert_slide_after(prs, src_slide):
    """src_slide の直後に、同じレイアウトの空スライドを挿入して返す。

    python-pptx には挿入位置を指定する API が無いため、末尾へ追加したうえで
    presentation の sldIdLst の並びを入れ替える。スライドの表示順は
    sldIdLst の順序で決まるため、これで直後の位置に置ける。
    """
    new_slide = prs.slides.add_slide(src_slide.slide_layout)
    # レイアウト由来のプレースホルダは中身が無いので取り除く
    for shp in list(new_slide.shapes):
        sp = shp._element
        sp.getparent().remove(sp)
    sldIdLst = prs.part._element.find("p:sldIdLst", NS)
    r_id = R_ATTR_PREFIX + "id"
    src_id = None
    for sldId in sldIdLst:
        if prs.part.rels[sldId.get(r_id)].target_part is src_slide.part:
            src_id = sldId
    if src_id is None:
        raise ValueError("元のスライドの sldId が見つかりません")
    new_id = sldIdLst[-1]  # add_slide で末尾に付いた分
    sldIdLst.remove(new_id)
    sldIdLst.insert(list(sldIdLst).index(src_id) + 1, new_id)
    return new_slide


def apply_splits(prs, slide_objs, slides, verdicts, cfg):
    """分割候補のスライドを2枚に分ける。変更の一覧を返す。

    slide_objs は分割前の prs.slides の並び（slides と対応）。
    元の番号がずれないよう、後ろのスライドから処理する。
    見出しは元のスライドに残したうえで複製を新しいスライドにも置き、
    後半の要素は元の位置のまま新しいスライドへ移す。
    """
    changes = []

    def rec(sno, type_, head, detail):
        changes.append(
            {
                "slide": sno,
                "type": type_,
                "elem": "（スライド全体）",
                "label": "",
                "head": head[:24].replace("\n", " "),
                "detail": detail,
            }
        )

    targets = [v for v in verdicts if v["candidate"]]
    for v in sorted(targets, key=lambda v: -v["slide"]):
        sno = v["slide"]
        elems = slides[sno - 1]
        src_slide = slide_objs[sno - 1]
        heading = _split_heading(elems, cfg)
        movable = _split_movable(elems, heading, cfg)
        reason = _split_reason(v, cfg)
        head = heading.text if heading is not None else ""
        if len(movable) < 2:
            rec(
                sno,
                "未解消",
                head,
                "分割候補だが、移せる要素が %d個しかないため分割しなかった。判定根拠: %s"
                % (len(movable), reason),
            )
            continue
        to_move = movable[len(movable) // 2 :]
        new_slide = _insert_slide_after(prs, src_slide)
        if heading is not None:
            dup = copy.deepcopy(heading.shape._element)
            _retarget_rels(src_slide.part, new_slide.part, dup)
            new_slide.shapes._spTree.append(dup)
        for e in to_move:
            el = e.shape._element
            _retarget_rels(src_slide.part, new_slide.part, el)
            new_slide.shapes._spTree.append(el)
        heading_note = (
            "見出し「%s」は元のスライドと新しいスライドの両方に置いた。"
            % head[:16].replace("\n", " ")
            if heading is not None
            else ""
        )
        rec(
            sno,
            "分割",
            head,
            "元のスライド%dを2枚に分け、後半の%d要素を直後の新しいスライドへ移した。%s判定根拠: %s"
            % (sno, len(to_move), heading_note, reason),
        )
    return changes


# ---------------------------------------------------------------- ページ並び替え（フェーズ4）
def _norm_for_match(text):
    """見出しと目次行の照合用にテキストを正規化する。

    空白をすべて取り除き、先頭の番号付け（「1.」「２、」など）と末尾の
    句点を落とす。番号や空白の有無で照合がぶれないようにするためで、
    照合はこの正規化後の一致・包含だけで行う。
    """
    t = re.sub(r"\s+", "", text)
    t = re.sub(r"^[0-9０-９]+[.．、:：)）]", "", t)
    return t.rstrip("。")


def _heading_of(elems, cfg):
    """スライドの見出し要素と、その先頭行のテキストを返す。無ければ (None, "")。

    見出しの決め方は分割（_split_heading）と同じ。テキストは先頭の
    空でない行だけを使う（複数行の見出しは1行目が表題という前提）。
    """
    heading = _split_heading(elems, cfg)
    if heading is None:
        return None, ""
    for line in heading.text.split("\n"):
        if line.strip():
            return heading, line.strip()
    return heading, ""


def _content_chars(elems, heading, cfg):
    """見出し要素と注記だけの要素を除いた、本文の文字数。

    「タイトルだけのスライド」の判定に使う。注記（出典・ページ番号など）は
    スライドの付属物として本文に数えない。注記だけの要素かの判定は、
    分割で元のスライドに残す判定に使った _is_note_only と同じ。
    """
    median = slide_body_median(elems, cfg)
    total = 0
    for e in elems:
        if e is heading or _is_note_only(e, cfg, median):
            continue
        total += sum(len(u.text) for u in e.units)
    return total


def _order_text(order, heads_text):
    """順序を「1「見出し」 → 2「見出し」」の形で返す。"""
    parts = []
    for no in order:
        head = heads_text.get(no, "")
        parts.append("%d「%s」" % (no, head[:16] if head else "（見出しなし）"))
    return " → ".join(parts)


def _title_candidate(sno, head_text, body_chars, slides, heads, cfg):
    """タイトルだけのスライド sno に対する候補。該当しなければ None。

    「内容に当たるスライド」は、見出し同士を正規化したうえで一方が他方を
    含む関係にあり、本文を持つ（タイトルだけではない）スライドとする。
    最も手前の内容スライドを控え、それより後ろにある場合だけ候補にする
    （既に前にあるなら直すところが無い）。
    """
    limit = cfg["reorder_title_only_max_body_chars"]
    nh = _norm_for_match(head_text)
    if not nh:
        return None
    anchor = None
    anchor_head = ""
    for j, elems in enumerate(slides, 1):
        if j == sno:
            continue
        h_elem, h_text = heads[j]
        if h_elem is None or not h_text:
            continue
        nj = _norm_for_match(h_text)
        if not nj or (nh not in nj and nj not in nh):
            continue
        if _content_chars(elems, h_elem, cfg) <= limit:
            continue  # 見出しが関係しても、相手もタイトルだけなら内容とは言えない
        anchor, anchor_head = j, h_text
        break
    if anchor is None or anchor > sno:
        return None
    order = list(range(1, len(slides) + 1))
    order.remove(sno)
    order.insert(anchor - 1, sno)
    nj = _norm_for_match(anchor_head)
    if nh == nj:
        rel = "見出しが一致する"
    elif nh in nj:
        rel = "見出し「%s」が相手の見出しに含まれる" % head_text
    else:
        rel = "相手の見出し「%s」がこの見出しに含まれる" % anchor_head
    reason = (
        "スライド%d「%s」は見出しのみで、本文は%d文字（しきい値%d文字以下）しかない。"
        "%sスライド%d「%s」が内容に当たり、その後ろにある"
        % (sno, head_text, body_chars, limit, rel, anchor, anchor_head)
    )
    return {
        "kind": "title",
        "slide": sno,
        "head": head_text,
        "order": order,
        "reason": reason,
    }


def _toc_candidate(sno, elems, heading, slides, heads, cfg):
    """目次に当たるスライド sno に対する候補。該当しなければ None。

    見出し要素と注記を除いた行のうち、他のスライドの見出しと正規化後に
    完全一致する行を目次の項目とみなす。一致が設定件数以上あり、項目の
    並び（= 対応するスライドの並び）と実際の順序が食い違う場合だけ候補に
    する。候補の順序は、食い違うスライド同士が占めている位置の中でだけ
    目次の並びに入れ替えたもので、関係ないスライドは動かさない。
    """
    median = slide_body_median(elems, cfg)
    lines = []
    for e in elems:
        if e is heading or _is_note_only(e, cfg, median):
            continue
        for u in e.units:
            for p in u.paragraphs:
                for line in p["text"].split("\n"):
                    if line.strip():
                        lines.append(line.strip())
    by_norm = {}
    for j, (h_elem, h_text) in heads.items():
        if j == sno or h_elem is None or not h_text:
            continue
        nj = _norm_for_match(h_text)
        if nj and nj not in by_norm:
            by_norm[nj] = j
    matches = []  # (行テキスト, 対応するスライド番号) を行の順に並べたもの
    seen = set()
    for line in lines:
        j = by_norm.get(_norm_for_match(line))
        if j is not None and j not in seen:
            seen.add(j)
            matches.append((line, j))
    if len(matches) < cfg["reorder_toc_min_matches"]:
        return None
    seq = [j for _line, j in matches]
    if seq == sorted(seq):
        return None  # 目次の並びと実際の順序が一致している
    order = list(range(1, len(slides) + 1))
    for k, pos in enumerate(sorted(seq)):
        order[pos - 1] = seq[k]
    detail = "、".join("「%s」→スライド%d" % (line, j) for line, j in matches)
    reason = (
        "スライド%dは目次に当たり、見出しと一致する行が%d件ある（%s）。"
        "行の並びはスライド%sの順だが、実際は%sの順に並んでいて食い違う"
        % (
            sno,
            len(matches),
            detail,
            "→".join(str(j) for j in seq),
            "→".join(str(j) for j in sorted(seq)),
        )
    )
    return {
        "kind": "toc",
        "slide": sno,
        "head": heads[sno][1],
        "order": order,
        "reason": reason,
    }


def judge_reorder(slides, cfg):
    """並び替え候補の一覧を返す。根拠がテキストから読み取れる場合だけ出す。

    候補は2種類で、どちらも現在の順序・候補の順序・根拠を持つ。
    番号（no）はレポートへの掲載順に振り、--reorder-apply はこの番号で選ぶ。
    - title: タイトルだけのスライドが、その内容に当たるスライドより後ろにある
    - toc: 目次に当たるスライドの行の順序と、実際のスライド順序が食い違う
    """
    heads = {}
    for sno, elems in enumerate(slides, 1):
        heads[sno] = _heading_of(elems, cfg)
    heads_text = {sno: text for sno, (_elem, text) in heads.items()}
    limit = cfg["reorder_title_only_max_body_chars"]
    candidates = []
    for sno, elems in enumerate(slides, 1):
        heading, head_text = heads[sno]
        if heading is None or not head_text:
            continue
        toc = _toc_candidate(sno, elems, heading, slides, heads, cfg)
        if toc is not None:
            candidates.append(toc)
        body_chars = _content_chars(elems, heading, cfg)
        if body_chars > limit:
            continue
        title = _title_candidate(sno, head_text, body_chars, slides, heads, cfg)
        if title is not None:
            candidates.append(title)
    for i, c in enumerate(candidates, 1):
        c["no"] = i
        c["current"] = list(range(1, len(slides) + 1))
        c["heads"] = heads_text
    return candidates


def apply_reorder(prs, candidate):
    """候補の順序を sldIdLst の並び替えで適用し、変更の記録を返す。

    スライド部品そのものは作り直さず、表示順を決める sldIdLst の並びだけを
    入れ替える（分割の _insert_slide_after と同じ考え方）。既存の sldId を
    候補の順に append し直すと、その順に並ぶ。
    """
    sldIdLst = prs.part._element.find("p:sldIdLst", NS)
    ids = list(sldIdLst)
    for idx in candidate["order"]:
        sldIdLst.append(ids[idx - 1])
    return {
        "slide": candidate["slide"],
        "type": "並び替え",
        "elem": "（スライド全体）",
        "label": "",
        "head": candidate["head"],
        "detail": "候補%dを適用。現在の順序 %s → 候補の順序 %s。根拠: %s"
        % (
            candidate["no"],
            _order_text(candidate["current"], candidate["heads"]),
            _order_text(candidate["order"], candidate["heads"]),
            candidate["reason"],
        ),
    }


# ---------------------------------------------------------------- 構成の指摘の修正（番号で選んで適用）
# 結果の status の値: "適用"（直した）/ "解消済み"（適用直前の再確認で条件が
# 消えていた）/ "未解消"（試したが解消できなかった）/ "対象外"（直さない種類）


def _structure_result(no, sno, f, status, detail):
    """構成の指摘1件分の修正結果の記録。"""
    return {
        "no": no,
        "slide": sno,
        "type": f["type"],
        "elem": f["elem"].name,
        "label": f["label"],
        "head": f["head"],
        "status": status,
        "detail": detail,
    }


def _refresh_unit_sizes(resolver, elem, unit):
    """読み取り済みの run のサイズを、現在の XML から取り直す。

    --fix や先に適用した構成の修正で run のサイズが変わっていると、読み取り
    時のサイズのままでは再確認や修正量の計算がずれる。取り除かれた段落の
    run は親を辿れないため、そのまま残す。
    """
    for p in unit.paragraphs:
        runs = []
        for (n, size, src), (el, _txt) in zip(p["runs"], p["run_els"]):
            p_el = el.getparent()
            if p_el is None:
                runs.append((n, size, src))
                continue
            new_size, new_src = resolver.resolve(elem.shape, p_el, el)
            runs.append((n, new_size, new_src))
        p["runs"] = runs


def apply_structure_fixes(prs, slides, cfg, selected, resolver):
    """番号で選ばれた構成の指摘を修正し、結果の記録を番号順で返す。

    selected は _structure_numbered の (番号, スライド番号, 指摘) の並び。
    適用の直前に、その指摘の条件がまだ成り立つかをスライドごとに再確認する
    （--fix と併用すると、サイズや位置が変わって条件が消えている場合がある）。
    条件が消えていれば変更せず、その旨を結果に記録する。
    """
    slide_objs = list(prs.slides)
    by_slide = {}
    for no, sno, f in selected:
        by_slide.setdefault(sno, []).append((no, f))
    results = []
    for sno in sorted(by_slide):
        elems = slides[sno - 1]
        for e in elems:
            for u in e.units:
                _refresh_unit_sizes(resolver, e, u)
        median = slide_body_median(elems, cfg)
        for no, f in sorted(by_slide[sno]):
            results.append(
                _apply_structure_one(
                    slide_objs[sno - 1],
                    sno,
                    elems,
                    cfg,
                    no,
                    f,
                    median,
                    prs.slide_width,
                )
            )
    results.sort(key=lambda r: r["no"])
    return results


def _apply_structure_one(slide, sno, elems, cfg, no, f, median, slide_w):
    """構成の指摘1件を修正する。種類ごとに処理を振り分ける。"""
    if f["type"] == "長い注釈":
        return _structure_result(
            no,
            sno,
            f,
            "対象外",
            "長い注釈は自動修正の対象外のため、変更しなかった。短くするか大きくするかは、内容を見て判断する",
        )
    if f["type"] == "改行の孤立":
        return _apply_structure_orphan(sno, elems, cfg, no, f, median, slide_w)
    if f["type"] == "色の不揃い":
        return _apply_structure_color(sno, elems, cfg, no, f)
    if f["type"] == "見出しの繰り返し":
        return _apply_structure_repeat(slide, sno, elems, cfg, no, f, median)
    return _structure_result(
        no,
        sno,
        f,
        "対象外",
        "この種類の指摘は自動修正に対応していないため、変更しなかった",
    )


def _orphan_paragraph(elem, cfg, median, label):
    """要素の中で、改行の孤立が起きている最初の (ユニット, 段落) を返す。無ければ None。

    判定（_structure_orphans）と同じ物差しで探す。対象は本文とラベルで、
    ラベル（表のセルなど）が指定された場合はそのユニットだけを見る。
    """
    max_chars = cfg.get("structure_orphan_max_chars", 3)
    min_chars = cfg.get("structure_orphan_min_chars", 12)
    for u in elem.units:
        if label and u.label != label:
            continue
        if not u.text.strip() or u.inner_w <= 0:
            continue
        role, _why = classify(elem, u, cfg, median)
        if role not in ("body", "label"):
            continue
        inner_w_pt = mm(u.inner_w) * FIT_SAFETY / PT_TO_MM
        for p in u.paragraphs:
            segs = _paragraph_segments(u, p)
            if segs is None:
                continue
            total = sum(len(s) for s in segs)
            if total < min_chars or not segs[-1]:
                continue
            counts = _wrapped_line_counts(segs[-1], inner_w_pt)
            if len(counts) >= 2 and counts[-1] <= max_chars:
                return u, p
    return None


def _orphan_resolved_at(seg_sizes, inner_w_pt, cfg):
    """最後の区切りを seg_sizes（(文字, サイズ) の並び）で折り返したとき、孤立しないかを返す。"""
    max_chars = cfg.get("structure_orphan_max_chars", 3)
    if not seg_sizes:
        return False
    counts = _wrapped_line_counts(seg_sizes, inner_w_pt)
    return len(counts) < 2 or counts[-1] > max_chars


def _widen_for_orphan(elems, e, unit, para, cfg, slide_w):
    """要素の幅を右へ広げて孤立を解消する。広げた後の幅(EMU)を返す。無理なら None。

    スライドの右端を超えず、元の幅に上限倍率を掛けた範囲まで、0.5mm 刻みで
    解消する最小の幅を探す。重なりの回避対象は、判定（judge）の重なりが
    数える相手と同じ範囲（テキストを持つ要素と、図(画像)・表・グラフ・
    SmartArt。線・コネクタは除く）で、新たに重ならない範囲に限る。
    表のセルは、要素の幅を広げてもセルの幅が変わらないため対象にしない。
    """
    if unit.label or e.shape is None or unit.inner_w <= 0:
        return None
    segs = _paragraph_segments(unit, para)
    if segs is None or not segs[-1]:
        return None
    max_ratio = cfg.get("structure_orphan_widen_max_ratio", 1.3)
    insets = e.w - unit.inner_w
    max_w = min(e.w * max_ratio, slide_w - e.x)
    if max_w <= e.w:
        return None
    others = [
        o
        for o in elems
        if o is not e
        and o.kind != "線・コネクタ"
        and (o.has_text or o.kind in ("図(画像)", "表", "グラフ", "SmartArt"))
    ]
    step = 0.5 * EMU_PER_MM
    new_w = e.w + step
    while new_w <= max_w + 1e-6:
        inner_pt = mm(new_w - insets) * FIT_SAFETY / PT_TO_MM
        if _orphan_resolved_at(segs[-1], inner_pt, cfg) and not _widen_overlaps(
            e, new_w, others
        ):
            delta = new_w - e.w
            e.shape.width = int(round(e.shape.width + delta / e.tf.sx))
            e.w = new_w
            unit.inner_w += delta
            return new_w
        new_w += step
    return None


def _widen_overlaps(e, new_w, others):
    """幅を new_w に広げたとき、others のいずれかと新たに重なるかを返す。"""
    for o in others:
        iy = min(e.y + e.h, o.y + o.h) - max(e.y, o.y)
        if iy <= 0:
            continue
        old_area = max(min(e.x + e.w, o.x + o.w) - max(e.x, o.x), 0) * iy
        new_area = max(min(e.x + new_w, o.x + o.w) - max(e.x, o.x), 0) * iy
        if new_area > old_area + 1e-6:
            return True
    return False


def _shrink_paragraph(unit, para, ratio):
    """段落内の run のサイズを比率で下げる（_apply_size の縮小版）。"""
    runs = []
    for (n, size, src), (el, _txt) in zip(para["runs"], para["run_els"]):
        if size is None:
            runs.append((n, size, src))
            continue
        raw = size * ratio
        rpr = el.find("a:rPr", NS)
        if rpr is None:
            rpr = el.makeelement(A + "rPr", {})
            el.insert(0, rpr)
        rpr.set("sz", str(int(round(raw * 100))))
        runs.append((n, raw, src))
    para["runs"] = runs


def _shrink_for_orphan(unit, para, cfg):
    """段落の文字を小さくして孤立を解消する。(元の最大サイズ, 変更後のサイズ) を返す。

    下げ幅は、段落の最後の区切りの最大サイズから 0.5pt 刻みで探す。下限は
    「最大で structure_orphan_shrink_max_pt pt まで下げる」と「元の9割まで」
    のうち大きい方。段落全体の run を同じ比率で下げる。
    """
    segs = _paragraph_segments(unit, para)
    if segs is None or not segs[-1]:
        return None
    rep = max(size for _ch, size in segs[-1])
    floor = max(rep - cfg.get("structure_orphan_shrink_max_pt", 2), rep * 0.9)
    inner_w_pt = mm(unit.inner_w) * FIT_SAFETY / PT_TO_MM
    target = rep - 0.5
    while target >= floor - 1e-9:
        ratio = target / rep
        scaled = [(ch, size * ratio) for ch, size in segs[-1]]
        if _orphan_resolved_at(scaled, inner_w_pt, cfg):
            _shrink_paragraph(unit, para, ratio)
            return rep, target
        target -= 0.5
    return None


def _apply_structure_orphan(sno, elems, cfg, no, f, median, slide_w):
    """改行の孤立を、幅の拡大→文字の縮小の順に試して解消する。"""
    e = f["elem"]
    still = any(
        f2["elem"] is e and f2["label"] == f["label"]
        for f2 in _structure_orphans(elems, cfg, median)
    )
    target = _orphan_paragraph(e, cfg, median, f["label"]) if still else None
    if target is None:
        return _structure_result(
            no,
            sno,
            f,
            "解消済み",
            "適用の直前に再確認したところ、孤立はすでに解消していたため変更しなかった",
        )
    unit, para = target
    old_w = e.w
    new_w = _widen_for_orphan(elems, e, unit, para, cfg, slide_w)
    if new_w is not None:
        return _structure_result(
            no,
            sno,
            f,
            "適用",
            "要素の幅を %.1fmm から %.1fmm へ広げて、最後の行の孤立を解消した"
            % (mm(old_w), mm(new_w)),
        )
    shrunk = _shrink_for_orphan(unit, para, cfg)
    if shrunk is not None:
        rep, target_size = shrunk
        return _structure_result(
            no,
            sno,
            f,
            "適用",
            "段落の文字を %.1fpt から %.1fpt へ小さくして、最後の行の孤立を解消した"
            % (rep, target_size),
        )
    return _structure_result(
        no,
        sno,
        f,
        "未解消",
        "幅の拡大（元の幅の最大 %g 倍まで・他の要素と新たに重ならない範囲）でも、文字の縮小（最大 %gpt・元の9割まで）でも、孤立を解消できなかった"
        % (
            cfg.get("structure_orphan_widen_max_ratio", 1.3),
            cfg.get("structure_orphan_shrink_max_pt", 2),
        ),
    )


def _color_majority_for(elems, cfg, target):
    """target を含む並びの組で、target 以外が共有する塗りの色を返す。無ければ None。

    判定（_structure_color_mismatch）と同じ物差し（_color_groups）で組を
    作り直し、target が1つだけ違う側にいる組を探す。
    """
    for _axis, group in _color_groups(elems, cfg):
        if not any(e is target for e, _c in group):
            continue
        others = [c for e, c in group if e is not target]
        target_color = next(c for e, c in group if e is target)
        if others and len(set(others)) == 1 and others[0] != target_color:
            return others[0]
    return None


def _apply_structure_color(sno, elems, cfg, no, f):
    """色の不揃いを、並んだ他の要素と同じ色に合わせて直す。"""
    e = f["elem"]
    still = any(f2["elem"] is e for f2 in _structure_color_mismatch(elems, cfg))
    if not still:
        return _structure_result(
            no,
            sno,
            f,
            "解消済み",
            "適用の直前に再確認したところ、色の不揃いはすでに解消していたため変更しなかった",
        )
    main_color = _color_majority_for(elems, cfg, e)
    if main_color is None:
        return _structure_result(
            no,
            sno,
            f,
            "対象外",
            "並びの組を確認できなかったため、変更しなかった",
        )
    if e.shape is None or e.kind in (
        "表",
        "グラフ",
        "SmartArt",
        "図(画像)",
        "線・コネクタ",
    ):
        return _structure_result(
            no,
            sno,
            f,
            "対象外",
            "この要素は塗りを単色で設定できないため、変更しなかった",
        )
    _kind, old_color = _fill_state(e)
    if old_color is None:
        return _structure_result(
            no,
            sno,
            f,
            "対象外",
            "現在の塗りの色を確認できなかったため、変更しなかった",
        )
    try:
        e.shape.fill.solid()
        e.shape.fill.fore_color.rgb = RGBColor(*main_color)
    except Exception:
        return _structure_result(
            no,
            sno,
            f,
            "対象外",
            "この要素は塗りを単色で設定できなかったため、変更しなかった",
        )
    return _structure_result(
        no,
        sno,
        f,
        "適用",
        "塗りの色を #%02X%02X%02X から、並んだ他の要素と同じ #%02X%02X%02X に合わせた"
        % (
            old_color[0],
            old_color[1],
            old_color[2],
            main_color[0],
            main_color[1],
            main_color[2],
        ),
    )


def _live_paragraphs(elem):
    """要素のテキスト枠の段落を、実 XML の (a:p, 正規化したテキスト) で返す。

    読み取り済みの段落は、先に適用した修正でずれている場合があるため、
    取り除く対象は実 XML で数え直す。テキスト枠が取れなければ None。
    """
    if elem.shape is None:
        return None
    txbody = elem.shape._element.find("p:txBody", NS)
    if txbody is None:
        return None
    out = []
    for p_el in txbody.findall("a:p", NS):
        text = "".join(t.text or "" for t in p_el.iter(A + "t"))
        out.append((p_el, _norm_short_text(text)))
    return out


def _apply_structure_repeat(slide, sno, elems, cfg, no, f, median):
    """見出しの繰り返しを、各要素から段落を取り除き、共通の見出しを1つ作って直す。

    取り除くのは、その文と一致する独立した段落だけ。共通の見出しの配置先が
    確保できない場合は、段落を取り除かずに対象外とする（配置できることを
    確認してから取り除く順序にする）。
    """
    t = f["head"]
    still = any(f2["head"] == t for f2 in _structure_repeats(elems, cfg, median))
    if not still:
        return _structure_result(
            no,
            sno,
            f,
            "解消済み",
            "適用の直前に再確認したところ、繰り返しはすでに解消していたため変更しなかった",
        )
    min_count = cfg.get("structure_repeat_min", 3)
    # 対象の要素: 注記でないユニットに、その文と一致する段落を持つもの
    targets = []  # (要素, 実XMLの段落一覧, 取り除く a:p, 段落の最大サイズ)
    for e in elems:
        if e.kind == "表":
            continue
        has_para = False
        max_size = None
        for u in e.units:
            if not u.text.strip():
                continue
            role, _why = classify(e, u, cfg, median)
            if role == "note":
                continue
            for p in u.paragraphs:
                if _norm_short_text(p["text"]) != t:
                    continue
                has_para = True
                for _n, size, _src in p["runs"]:
                    if size is not None:
                        eff = size * u.scale
                        max_size = eff if max_size is None else max(max_size, eff)
        if not has_para:
            continue
        live = _live_paragraphs(e)
        if live is None:
            continue
        matched = [p_el for p_el, text in live if text == t]
        if not matched:
            continue
        targets.append((e, live, matched, max_size))
    if len(targets) < min_count:
        return _structure_result(
            no,
            sno,
            f,
            "対象外",
            "取り除ける段落を持つ要素が %d 個しか確認できなかったため、変更しなかった"
            % len(targets),
        )
    # 共通の見出しの配置先を、段落を取り除く前に確認する
    sizes = [ms for _e, _l, _m, ms in targets if ms is not None]
    size_pt = max(sizes) if sizes else float(cfg.get("label_min_pt", 18))
    head_h = int(size_pt * LINE_HEIGHT * PT_TO_MM * EMU_PER_MM) + DEFAULT_INSET_TB * 2
    gap = 2.0 * EMU_PER_MM
    min_x = min(e.x for e, _l, _m, _s in targets)
    min_y = min(e.y for e, _l, _m, _s in targets)
    max_r = max(e.x + e.w for e, _l, _m, _s in targets)
    top = min_y - gap - head_h
    if top < 0:
        return _structure_result(
            no,
            sno,
            f,
            "対象外",
            "共通の見出しを置く余白がスライドの上端に無いため、段落は取り除かず変更しなかった",
        )
    target_ids = {id(e) for e, _l, _m, _s in targets}
    for o in elems:
        if id(o) in target_ids or not o.has_text:
            continue
        ix = min(max_r, o.x + o.w) - max(min_x, o.x)
        iy = min(top + head_h, o.y + o.h) - max(top, o.y)
        if ix > EMU_PER_MM and iy > EMU_PER_MM:
            return _structure_result(
                no,
                sno,
                f,
                "対象外",
                "共通の見出しの配置先に他の要素（「%s」）が重なるため、段落は取り除かず変更しなかった"
                % (o.text[:16].replace("\n", " ")),
            )
    # 段落を取り除く。テキスト枠の唯一の段落なら、段落は残して中身だけ空にする
    for _e, live, matched, _s in targets:
        if len(live) - len(matched) == 0:
            for p_el in matched:
                for child in list(p_el):
                    if child.tag in (A + "r", A + "fld", A + "br"):
                        p_el.remove(child)
        else:
            for p_el in matched:
                p_el.getparent().remove(p_el)
    # 共通の見出しを作る
    txbox = slide.shapes.add_textbox(int(min_x), int(top), int(max_r - min_x), head_h)
    tf = txbox.text_frame
    tf.word_wrap = True
    tf.margin_left = DEFAULT_INSET_LR
    tf.margin_right = DEFAULT_INSET_LR
    tf.margin_top = DEFAULT_INSET_TB
    tf.margin_bottom = DEFAULT_INSET_TB
    run = tf.paragraphs[0].add_run()
    run.text = t
    run.font.size = Pt(size_pt)
    return _structure_result(
        no,
        sno,
        f,
        "適用",
        "「%s」の段落を %d 個の要素から取り除き、共通の見出しを1つ作って組の上に置いた（文字の大きさ %.1fpt）"
        % (t, len(targets), size_pt),
    )


# ---------------------------------------------------------------- バージョン更新の確認
def _version_key(version):
    """バージョン文字列を数値のタプルにする。解釈できなければ None。

    先頭の v は外し、"1.10.0" と "1.9.0" を文字列ではなく数値で比べる。
    桁数は3つにそろえる（"1.2" と "1.2.0" を同じとみなす）。
    """
    s = str(version).strip()
    if s[:1] in ("v", "V"):
        s = s[1:]
    if not re.search(r"\d", s):
        return None
    nums = []
    for part in s.split("."):
        m = re.match(r"\d+", part)
        nums.append(int(m.group()) if m else 0)
    return tuple((nums + [0, 0, 0])[:3])


def is_dev_version(version):
    """開発版かどうか（-dev を含む、または 0.0.0）。"""
    return "-dev" in version or version.strip() == "0.0.0"


def should_auto_check(version=None):
    """GUI 起動時に自動で更新確認をするか。開発版ではしない。"""
    return not is_dev_version(__version__ if version is None else version)


def pick_release_asset(assets):
    """リリースの asset 一覧から、配布用 zip のダウンロード URL を選ぶ。

    名前が「ReSlide-」で始まり「.zip」で終わるもの。無ければ None。
    """
    for a in assets or []:
        name = a.get("name") or ""
        if name.startswith("ReSlide-") and name.endswith(".zip"):
            return a.get("browser_download_url")
    return None


def interpret_update_response(data, current_version):
    """GitHub API の応答（releases/latest の JSON）を、更新確認の結果に直す。

    ネットワークには触らない。結果は status（update_available / up_to_date /
    failed）、latest_tag、zip_url、html_url、current を持つ dict。
    """
    result = {
        "status": "failed",
        "latest_tag": None,
        "zip_url": None,
        "html_url": UPDATE_RELEASES_URL,
        "current": current_version,
    }
    if not isinstance(data, dict):
        return result
    tag = data.get("tag_name")
    latest_key = _version_key(tag) if tag else None
    current_key = _version_key(current_version)
    if latest_key is None or current_key is None:
        return result
    result["latest_tag"] = tag
    if data.get("html_url"):
        result["html_url"] = data["html_url"]
    if latest_key > current_key:
        result["status"] = "update_available"
        result["zip_url"] = pick_release_asset(data.get("assets"))
    else:
        result["status"] = "up_to_date"
    return result


def check_update(current_version=None):
    """GitHub の最新リリースを確認し、結果の dict を返す。

    失敗しても例外は投げず、status="failed" で返す（オフラインでも
    呼び出し側が止まらないように）。current_version はテスト用に差し替えられる。
    """
    current = __version__ if current_version is None else current_version
    failed = {
        "status": "failed",
        "latest_tag": None,
        "zip_url": None,
        "html_url": UPDATE_RELEASES_URL,
        "current": current,
    }
    req = urllib.request.Request(
        UPDATE_API_URL,
        headers={
            "User-Agent": UPDATE_USER_AGENT,
            "Accept": "application/vnd.github+json",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=UPDATE_TIMEOUT_SEC) as res:
            data = json.loads(res.read().decode("utf-8"))
    except Exception:
        return failed
    return interpret_update_response(data, current)


def format_update_check_line(result):
    """--check-update の標準出力（1〜2行）を返す。"""
    if result["status"] == "update_available":
        return "新しいバージョン %s があります（現在のバージョン v%s）\nダウンロード: %s" % (
            result["latest_tag"],
            result["current"],
            result["zip_url"] or result["html_url"],
        )
    if result["status"] == "up_to_date":
        return "最新です（現在のバージョン v%s）" % result["current"]
    return "更新の確認に失敗しました"


def update_banner_info(result):
    """更新確認の結果から、GUI のバナーに出す内容を決める。不要なら None。

    zip の URL があればボタンを「ダウンロード」にし、無ければ
    「Releasesを開く」にする。tkinter には触らない（テスト用に分けてある）。
    """
    if not result or result.get("status") != "update_available":
        return None
    zip_url = result.get("zip_url")
    return {
        "text": "新しいバージョン %s があります" % result.get("latest_tag"),
        "latest_tag": result.get("latest_tag"),
        "zip_url": zip_url,
        "releases_url": result.get("html_url") or UPDATE_RELEASES_URL,
        "download_label": "ダウンロード" if zip_url else "Releasesを開く",
    }


def download_update_zip(url, dest_dir=None, progress=None):
    """更新の zip を一時フォルダへダウンロードし、保存先の Path を返す。

    保存先は一時フォルダの下の ReSlideUpdate フォルダ（dest_dir で変更可）。
    progress は 0〜100 の整数（全体サイズが分からない時は None）で呼ばれる。
    差し替えはせず、保存までで止める（開く・入れ替えは利用者の手に委ねる）。
    """
    folder = (
        Path(dest_dir)
        if dest_dir
        else Path(tempfile.gettempdir()) / "ReSlideUpdate"
    )
    folder.mkdir(parents=True, exist_ok=True)
    name = url.split("/")[-1].split("?")[0] or "ReSlide-update.zip"
    dest = folder / name
    req = urllib.request.Request(url, headers={"User-Agent": UPDATE_USER_AGENT})
    with urllib.request.urlopen(req, timeout=UPDATE_TIMEOUT_SEC) as res:
        total = res.headers.get("Content-Length")
        total = int(total) if total and total.isdigit() else None
        done = 0
        with open(dest, "wb") as f:
            while True:
                chunk = res.read(65536)
                if not chunk:
                    break
                f.write(chunk)
                done += len(chunk)
                if progress is not None:
                    progress(int(done * 100 / total) if total else None)
    return dest


# ---------------------------------------------------------------- GUI（--gui）
def build_structure_rows(structure_findings):
    """GUI の一覧に表示する行データを、レポートと同じ番号・順序で返す。

    修正できるか（fixable）は種類で決まる。長い注釈は自動修正の対象外。
    """
    rows = []
    for no, sno, f in _structure_numbered(structure_findings):
        rows.append(
            {
                "no": no,
                "slide": sno,
                "type": f["type"],
                "head": f["head"],
                "detail": f["detail"],
                "fixable": f["type"] != "長い注釈",
            }
        )
    return rows


def apply_gui_selection(
    prs, slides, cfg, resolver, structure_findings, checked_numbers, do_fix
):
    """GUI で選ばれた構成の指摘を修正する。(変更一覧, 構成の修正結果) を返す。

    CLI の --structure-apply と同じ順序で、自動調整（任意）を先に行い、
    その後にチェックされた構成の修正を適用する。修正の中身は
    apply_structure_fixes を CLI と共有する。
    """
    changes = []
    if do_fix:
        changes = apply_fixes(slides, cfg, prs.slide_width, prs.slide_height)
    wanted = set(checked_numbers)
    selected = [
        item for item in _structure_numbered(structure_findings) if item[0] in wanted
    ]
    results = (
        apply_structure_fixes(prs, slides, cfg, selected, resolver) if selected else []
    )
    return changes, results


def save_gui_result(src, prs, changes, structure_results):
    """GUI の保存先。元のファイルと同じフォルダへ _reslide.pptx と変更一覧を書く。"""
    out_pptx = src.with_name(src.stem + "_reslide.pptx")
    prs.save(str(out_pptx))
    report = build_changes_report(src, out_pptx.name, changes, structure_results)
    out_md = src.with_name(src.stem + "_reslide_changes.md")
    out_md.write_text(report, encoding="utf-8")
    return out_pptx, out_md


def structure_results_summary(results):
    """構成の修正結果の要約（件数と1件ごとの結果）をテキストで返す。GUI の表示用。"""
    if not results:
        return "構成の指摘の修正: 指定なし"
    n_ok = sum(1 for r in results if r["status"] == "適用")
    n_gone = sum(1 for r in results if r["status"] == "解消済み")
    n_fail = sum(1 for r in results if r["status"] == "未解消")
    n_na = sum(1 for r in results if r["status"] == "対象外")
    L = [
        "構成の指摘の修正: %d件を指定（修正した %d件 / すでに解消していた %d件 / 解消できなかった %d件 / 対象外 %d件）"
        % (len(results), n_ok, n_gone, n_fail, n_na)
    ]
    for r in results:
        L.append(
            "- 番号%d [%s] スライド%d: %s"
            % (r["no"], r["type"], r["slide"], r["detail"])
        )
    return "\n".join(L)


def run_gui(src, prs, slides, cfg, resolver):
    """構成の指摘を一覧で見て、選んで修正するウィンドウを開く。

    tkinter の import はこの関数の中だけに閉じる（表示環境が無い場所でも
    CLI が壊れないように）。一覧データの組み立てと適用処理は、表示なしで
    テストできる関数（build_structure_rows / apply_gui_selection）に分けてある。
    """
    import tkinter as tk
    from tkinter import messagebox

    structure_findings = judge_structure(slides, cfg, prs.slide_width, prs.slide_height)
    rows = build_structure_rows(structure_findings)

    root = tk.Tk()
    root.title("ReSlide - %s" % src.name)
    root.geometry("900x680")

    tk.Label(root, text="対象ファイル: %s" % src.name, anchor="w").pack(
        fill="x", padx=10, pady=(10, 2)
    )
    count_label = tk.Label(
        root,
        text="構成の指摘 %d件。直したいものにチェックを入れてください。" % len(rows),
        anchor="w",
    )
    count_label.pack(fill="x", padx=10, pady=(0, 6))

    outer = tk.Frame(root)
    outer.pack(fill="both", expand=True, padx=10)
    canvas = tk.Canvas(outer)
    scrollbar = tk.Scrollbar(outer, orient="vertical", command=canvas.yview)
    inner = tk.Frame(canvas)
    inner.bind(
        "<Configure>", lambda _ev: canvas.configure(scrollregion=canvas.bbox("all"))
    )
    canvas.create_window((0, 0), window=inner, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    check_vars = []
    for row in rows:
        var = tk.BooleanVar(value=False)
        check_vars.append(var)
        frame = tk.Frame(inner)
        frame.pack(fill="x", anchor="w", pady=2)
        label = "番号%d スライド%d [%s] 「%s」" % (
            row["no"],
            row["slide"],
            row["type"],
            row["head"],
        )
        if row["fixable"]:
            cb = tk.Checkbutton(frame, text=label, variable=var, anchor="w")
        else:
            cb = tk.Checkbutton(
                frame,
                text=label + "（自動修正の対象外）",
                variable=var,
                state="disabled",
                anchor="w",
            )
        cb.pack(fill="x", anchor="w")
        tk.Label(
            frame,
            text=row["detail"],
            wraplength=840,
            justify="left",
            anchor="w",
            fg="#444444",
        ).pack(fill="x", padx=(24, 0), anchor="w")
    if not rows:
        tk.Label(inner, text="構成の指摘はありません。", anchor="w").pack(fill="x")

    fix_var = tk.BooleanVar(value=True)
    fix_check = tk.Checkbutton(
        root, text="文字の大きさや配置の自動調整も一緒に行う", variable=fix_var
    )
    fix_check.pack(anchor="w", padx=10, pady=(8, 0))

    result_text = tk.Text(root, height=9, wrap="word")
    result_text.pack(fill="x", padx=10, pady=8)
    result_text.insert("1.0", "まだ保存していません。")
    result_text.configure(state="disabled")

    def show_result(text):
        result_text.configure(state="normal")
        result_text.delete("1.0", "end")
        result_text.insert("1.0", text)
        result_text.configure(state="disabled")

    def on_apply():
        checked = [row["no"] for row, var in zip(rows, check_vars) if var.get()]
        if not checked and not fix_var.get():
            show_result("修正する指摘が選ばれていません。")
            return
        changes, results = apply_gui_selection(
            prs,
            slides,
            cfg,
            resolver,
            structure_findings,
            checked,
            fix_var.get(),
        )
        out_pptx, out_md = save_gui_result(src, prs, changes, results)
        show_result(
            "\n".join(
                [
                    "保存しました。",
                    "保存先: %s" % out_pptx,
                    "変更一覧: %s" % out_md,
                    "自動調整の変更: %d件" % len(changes),
                    structure_results_summary(results),
                ]
            )
        )
        apply_button.configure(state="disabled")
        fix_check.configure(state="disabled")

    buttons = tk.Frame(root)
    buttons.pack(fill="x", padx=10, pady=(0, 10))
    apply_button = tk.Button(
        buttons, text="選択した指摘を修正して保存", command=on_apply
    )
    apply_button.pack(side="left")
    close_button = tk.Button(buttons, text="閉じる", command=root.destroy)
    close_button.pack(side="left", padx=8)

    # --- バージョン更新の確認（起動時は裏で確認し、あれば上部にバナーを出す）
    update_state = {"running": False}
    banner = tk.Frame(root, bg="#fdf3d8")
    banner_label = tk.Label(banner, text="", bg="#fdf3d8", anchor="w")
    banner_label.pack(side="left", padx=(8, 4), pady=4)
    download_button = tk.Button(banner, text="ダウンロード")
    download_button.pack(side="left", padx=2, pady=4)
    tk.Button(banner, text="後で", command=banner.pack_forget).pack(
        side="left", padx=2, pady=4
    )

    def open_saved_folder(folder):
        if hasattr(os, "startfile"):
            os.startfile(str(folder))  # Windows のエクスプローラーで開く
        else:
            messagebox.showinfo("更新のダウンロード", "保存先: %s" % folder)

    def on_download(info):
        download_button.configure(state="disabled", text="ダウンロード中… 0%")

        def on_progress(pct):
            if pct is None:
                download_button.configure(text="ダウンロード中…")
            else:
                download_button.configure(text="ダウンロード中… %d%%" % pct)

        def on_done(saved):
            download_button.configure(
                state="normal",
                text="保存先フォルダを開く",
                command=lambda: open_saved_folder(saved.parent),
            )
            show_result("新しいバージョンの zip を保存しました: %s" % saved)

        def on_failed():
            download_button.configure(
                state="normal",
                text="ダウンロード",
                command=lambda: on_download(info),
            )
            messagebox.showerror("更新のダウンロード", "ダウンロードに失敗しました")

        def work():
            try:
                saved = download_update_zip(
                    info["zip_url"],
                    progress=lambda pct: root.after(0, on_progress, pct),
                )
            except Exception:
                root.after(0, on_failed)
                return
            root.after(0, on_done, saved)

        threading.Thread(target=work, daemon=True).start()

    def show_update_banner(result):
        info = update_banner_info(result)
        if info is None:
            return
        banner_label.configure(text=info["text"])
        if info["zip_url"]:
            download_button.configure(
                state="normal",
                text="ダウンロード",
                command=lambda: on_download(info),
            )
        else:
            download_button.configure(
                state="normal",
                text="Releasesを開く",
                command=lambda: webbrowser.open(info["releases_url"]),
            )
        if not banner.winfo_ismapped():
            banner.pack(fill="x", padx=10, pady=(0, 6), before=count_label)

    def on_check_done(result, manual):
        update_state["running"] = False
        if result["status"] == "update_available":
            show_update_banner(result)
        elif manual:
            if result["status"] == "up_to_date":
                messagebox.showinfo(
                    "更新の確認", "最新です（現在のバージョン v%s）" % __version__
                )
            else:
                messagebox.showwarning("更新の確認", "確認に失敗しました")

    def start_update_check(manual):
        if update_state["running"]:
            if manual:
                messagebox.showinfo("更新の確認", "すでに確認中です")
            return
        update_state["running"] = True

        def work():
            result = check_update()
            root.after(0, on_check_done, result, manual)

        threading.Thread(target=work, daemon=True).start()

    tk.Button(
        buttons, text="更新を確認", command=lambda: start_update_check(True)
    ).pack(side="left", padx=8, before=close_button)

    if should_auto_check(__version__):
        start_update_check(False)

    root.mainloop()
    return 0


CANDIDATE_KIND_JP = {
    "title": "タイトルだけのスライドを、その内容の前へ出す",
    "toc": "目次の並びに合わせて入れ替える",
}


# ---------------------------------------------------------------- レポート
def build_report(
    path,
    prs,
    slides,
    findings,
    stats,
    cfg,
    split_verdicts=None,
    reorder_candidates=None,
    structure_findings=None,
):
    L = []
    L.append("# ReSlide レポート: %s" % Path(path).name)
    L.append("")
    L.append(
        "- スライド数: %d / サイズ: %.0f x %.0f mm（単位はすべて mm、フォントは pt）"
        % (len(slides), mm(prs.slide_width), mm(prs.slide_height))
    )
    L.append(
        "- 基準値: 本文の下限 %gpt / ラベルの下限 %gpt（設定ファイルの値）"
        % (cfg["body_min_pt"], cfg.get("label_min_pt", 18))
    )
    total = sum(len(v) for v in findings.values())
    L.append("- 指摘の合計: %d 件" % total)
    body_median_all = stats.get("body_median_all")
    if body_median_all is not None and body_median_all < cfg["body_min_pt"]:
        L.append("")
        L.append("## 資料全体の文字の小ささ")
        L.append("")
        L.append(
            "- 資料全体の文字が小さめです（本文の中央値 %.1fpt / 本文の下限 %gpt）"
            % (body_median_all, cfg["body_min_pt"])
        )
    L.append("")
    L.append("## フォントサイズの解決状況")
    L.append("")
    L.append("- 文字数（run単位で数えた全テキスト）: %d" % stats["chars_total"])
    L.append("- **未解決の文字数: %d**" % stats["chars_unresolved"])
    for src, n in sorted(stats["by_src"].items(), key=lambda x: -x[1]):
        mark = "（指定がなく、既定値で決まった分）" if src in DEFAULT_SOURCES else ""
        L.append("  - 解決元 %s: %d文字%s" % (src, n, mark))
    if stats["unresolved"]:
        L.append("")
        L.append("未解決の箇所（サイズが取れず、判定していない）:")
        for s, name, label, head in stats["unresolved"][:50]:
            L.append("  - スライド%d「%s」%s 冒頭「%s」" % (s, name, label, head))
    L.append("")
    L.append("## スライド別の指摘")
    for sno in range(1, len(slides) + 1):
        L.append("")
        L.append("### スライド %d（指摘 %d 件）" % (sno, len(findings[sno])))
        if not findings[sno]:
            L.append("")
            L.append("指摘なし")
        for f in findings[sno]:
            e = f["elem"]
            L.append(
                "- [%s] %s%s 「%s」"
                % (
                    f["type"],
                    e.kind,
                    (" " + f["label"]) if f["label"] else "",
                    f["head"],
                )
            )
            L.append(
                "  - 要素: %s / 位置(%.1f, %.1f)mm 大きさ %.1fx%.1fmm"
                % (e.name, mm(e.x), mm(e.y), mm(e.w), mm(e.h))
            )
            L.append("  - %s" % f["detail"])
    if structure_findings is not None:
        L.append("")
        L.append("## 構成の指摘（番号で選んで修正できる）")
        L.append("")
        s_total = sum(len(v) for v in structure_findings.values())
        L.append("- 合計: %d 件" % s_total)
        if not s_total:
            L.append("")
            L.append("なし")
        else:
            L.append(
                "- 修正するには --structure-apply に番号を指定する（カンマ区切りで複数可。例: --structure-apply 1,3）。長い注釈は自動修正の対象外"
            )
        numbered = {
            id(f): no for no, _sno, f in _structure_numbered(structure_findings)
        }
        for sno in range(1, len(slides) + 1):
            if not structure_findings[sno]:
                continue
            L.append("")
            L.append("### スライド %d" % sno)
            for f in structure_findings[sno]:
                e = f["elem"]
                L.append(
                    "- 番号%d [%s] %s%s 「%s」"
                    % (
                        numbered[id(f)],
                        f["type"],
                        e.kind,
                        (" " + f["label"]) if f["label"] else "",
                        f["head"],
                    )
                )
                L.append(
                    "  - 要素: %s / 位置(%.1f, %.1f)mm 大きさ %.1fx%.1fmm"
                    % (e.name, mm(e.x), mm(e.y), mm(e.w), mm(e.h))
                )
                L.append("  - %s" % f["detail"])
    L.append("")
    L.append("## 注記類として本文の下限を当てなかったもの")
    L.append("")
    if not stats["notes"]:
        L.append("なし")
    for s, name, label, why, size, head in stats["notes"]:
        L.append(
            "- スライド%d「%s」%s 「%s」 サイズ %s / 根拠: %s"
            % (
                s,
                name,
                label,
                head.replace("\n", " "),
                ("%.1fpt" % size) if size else "未解決",
                why,
            )
        )
    L.append("")
    L.append("## ラベルとして本文より低い下限を当てたもの")
    L.append("")
    if not stats["labels"]:
        L.append("なし")
    for s, name, label, why, size, head in stats["labels"]:
        L.append(
            "- スライド%d「%s」%s 「%s」 サイズ %s / 根拠: %s"
            % (
                s,
                name,
                label,
                head.replace("\n", " "),
                ("%.1fpt" % size) if size else "未解決",
                why,
            )
        )
    if split_verdicts is not None:
        L.append("")
        L.append("## ページ分割の候補（--split の判定）")
        L.append("")
        L.append(
            "- 候補の条件: テキスト要素 %d個以上 / 総文字数 %d以上 / テキスト同士の重なり %d件以上（3条件すべて）"
            % (
                cfg["split_min_text_elements"],
                cfg["split_min_chars"],
                cfg["split_min_overlaps"],
            )
        )
        for v in split_verdicts:
            L.append("- スライド%d: %s" % (v["slide"], _split_reason(v, cfg)))
    if reorder_candidates is not None:
        L.append("")
        L.append("## ページ並び替えの候補（--reorder の判定）")
        L.append("")
        if not reorder_candidates:
            L.append("候補なし")
        else:
            L.append("- 実行するには --reorder-apply <候補番号> を指定する")
            for c in reorder_candidates:
                L.append("")
                L.append("### 候補 %d: %s" % (c["no"], CANDIDATE_KIND_JP[c["kind"]]))
                L.append("- 現在の順序: %s" % _order_text(c["current"], c["heads"]))
                L.append("- 候補の順序: %s" % _order_text(c["order"], c["heads"]))
                L.append("- 根拠: %s" % c["reason"])
    L.append("")
    L.append("## 対象外・前提")
    L.append("")
    L.append("- グラフ・SmartArt の内部の文字は読まない（要素の位置と大きさのみ）。")
    L.append("- 回転した要素は回転前の矩形で判定する。")
    L.append(
        "- レイアウトやマスターから継承する値は、プレースホルダの対応付けが取れた範囲で解決する。表スタイル由来のサイズは見ない。"
    )
    L.append(
        "- 重なりは、テキストを持つ要素と、図・表・グラフ同士の組み合わせのみ判定（背景用の文字なし図形と線、図(画像)同士は除外）。"
    )
    L.append(
        "- 文字色のコントラストは、文字色が明示の RGB で取れる場合だけ判定する。背景は、文字が乗っている図形や表のセルの単色の塗りを先に探し、次にスライド・レイアウト・マスターの背景を探し、背景の定義が無ければ表示に合わせて白として判定する。背景のテーマ色参照はテーマの色定義から解決する。テーマ色や自動色の文字と、画像やグラデーションなど色が決められないものの上の文字は対象外。"
    )
    if structure_findings is not None:
        L.append(
            "- 構成の指摘（見出しの繰り返し・改行の孤立・色の不揃い・長い注釈）は、形のパターンで見分けたもの。通常の自動調整（--fix）では直さず、--structure-apply で番号を指定したものだけを修正する。長い注釈は自動修正の対象外。改行の孤立は文字幅の近似による推定で、実際の折り返しとずれる場合がある。"
        )
    return "\n".join(L) + "\n"


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="pptx の見やすさ問題をレポートする（--fix で自動調整、--split でページ分割、--reorder-apply で並び替えも行う）"
    )
    ap.add_argument("pptx")
    ap.add_argument(
        "--config", default=str(Path(__file__).with_name("reslide_config.json"))
    )
    ap.add_argument(
        "--out",
        default=None,
        help="Markdownの出力先（既定: カレントに <名前>_reslide_report.md、--fix または --reorder-apply 時は <名前>_reslide_changes.md）",
    )
    ap.add_argument(
        "--fix",
        action="store_true",
        help="自動調整して <元名>_reslide.pptx を新規保存し、変更一覧を出す",
    )
    ap.add_argument(
        "--split",
        action="store_true",
        help="詰め込みスライドの分割候補を判定する（--fix と併用すると分割も実行する）",
    )
    ap.add_argument(
        "--reorder",
        action="store_true",
        help="並び替え候補を判定してレポートに出す（自動では並び替えない）",
    )
    ap.add_argument(
        "--reorder-apply",
        type=int,
        default=None,
        metavar="候補番号",
        help="--reorder で出た候補のうち、指定した番号だけを並び替えて <元名>_reslide.pptx へ保存する",
    )
    ap.add_argument(
        "--structure-apply",
        type=str,
        default=None,
        metavar="番号",
        help="レポートの構成の指摘のうち、指定した番号（カンマ区切り）だけを修正して <元名>_reslide.pptx へ保存する",
    )
    ap.add_argument(
        "--gui",
        action="store_true",
        help="構成の指摘を一覧で見て選んで修正するウィンドウを開く",
    )
    ap.add_argument(
        "--check-update",
        action="store_true",
        help="新しいバージョンがあるかだけを確認して終わる（レポートは作らない）",
    )
    ap.add_argument(
        "--version", action="version", version="ReSlide %s" % __version__
    )
    args = ap.parse_args(argv)
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8")
        except Exception:
            pass
    if args.check_update:
        print(format_update_check_line(check_update()))
        return 0
    structure_nums = None
    if args.structure_apply is not None:
        try:
            structure_nums = sorted(
                {int(x) for x in args.structure_apply.split(",") if x.strip()}
            )
        except ValueError:
            structure_nums = []
        if not structure_nums:
            print(
                "--structure-apply の番号の指定が正しくありません（例: --structure-apply 1,3）",
                file=sys.stderr,
            )
            return 2
    if args.reorder_apply is not None and args.split:
        print(
            "--reorder-apply と --split は併用できません。分割でスライドが増えると、候補番号が指す並びが変わるためです",
            file=sys.stderr,
        )
        return 2
    if structure_nums is not None and args.split:
        print(
            "--structure-apply と --split は併用できません。分割でスライドが増えると、番号が指す指摘が変わるためです",
            file=sys.stderr,
        )
        return 2
    if structure_nums is not None and args.reorder_apply is not None:
        print(
            "--structure-apply と --reorder-apply は併用できません",
            file=sys.stderr,
        )
        return 2
    if args.reorder and args.fix and args.reorder_apply is None:
        print(
            "--reorder はレポート表示（--fix なし）で使ってください。並び替えの実行は --reorder-apply <候補番号> で指定します",
            file=sys.stderr,
        )
        return 2
    src = Path(args.pptx)
    cfg = json.loads(Path(args.config).read_text(encoding="utf-8"))
    before = sha256(src)
    prs = Presentation(str(src))
    resolver = Resolver(prs)
    slides = read_presentation(prs, resolver)
    if args.gui:
        return run_gui(src, prs, slides, cfg, resolver)
    # 分割の判定は、調整で位置やサイズが変わる前の状態に対して行う
    verdicts = judge_split(slides, cfg) if args.split else None
    # 並び替えの判定も調整前の状態に対して行う（候補番号は入力時点の順序で決まる）
    candidates = (
        judge_reorder(slides, cfg)
        if (args.reorder or args.reorder_apply is not None)
        else None
    )
    chosen = None
    if args.reorder_apply is not None:
        if not candidates:
            print(
                "並び替え候補が無いため、--reorder-apply %d は実行できません"
                % args.reorder_apply,
                file=sys.stderr,
            )
            return 2
        if not 1 <= args.reorder_apply <= len(candidates):
            print(
                "--reorder-apply の候補番号が範囲外です（候補は %d 件）"
                % len(candidates),
                file=sys.stderr,
            )
            return 2
        chosen = candidates[args.reorder_apply - 1]
    # 構成の指摘の番号は、調整前の状態に対して決める（レポートの番号と一致させる）
    structure_findings = None
    structure_selected = []
    if structure_nums is not None:
        structure_findings = judge_structure(
            slides, cfg, prs.slide_width, prs.slide_height
        )
        numbered = _structure_numbered(structure_findings)
        if not numbered:
            print(
                "構成の指摘が無いため、--structure-apply は実行できません",
                file=sys.stderr,
            )
            return 2
        for n in structure_nums:
            if not 1 <= n <= len(numbered):
                print(
                    "--structure-apply の番号が範囲外です（構成の指摘は %d 件）"
                    % len(numbered),
                    file=sys.stderr,
                )
                return 2
        wanted = set(structure_nums)
        structure_selected = [item for item in numbered if item[0] in wanted]
    if args.fix or chosen is not None or structure_selected:
        changes = []
        structure_results = None
        if args.fix:
            slide_objs = list(prs.slides)
            changes = apply_fixes(slides, cfg, prs.slide_width, prs.slide_height)
            if args.split:
                changes += apply_splits(prs, slide_objs, slides, verdicts, cfg)
        if structure_selected:
            # 構成の修正は自動調整の後に行う。適用直前に条件を再確認するため、
            # --fix で状態が変わった指摘は対象外として記録される
            structure_results = apply_structure_fixes(
                prs, slides, cfg, structure_selected, resolver
            )
        if chosen is not None:
            # 並び替えは最後に行う。--fix の調整はスライドの枚数も順序も変えないため、
            # 候補が指す番号は sldIdLst の並びとそのまま対応する
            changes.append(apply_reorder(prs, chosen))
        out_pptx = Path.cwd() / (src.stem + "_reslide.pptx")
        prs.save(str(out_pptx))
        report = build_changes_report(src, out_pptx.name, changes, structure_results)
        out = (
            Path(args.out)
            if args.out
            else Path.cwd() / (src.stem + "_reslide_changes.md")
        )
        out.write_text(report, encoding="utf-8")
        print(report)
        print("調整後ファイル: %s" % out_pptx.resolve())
        print("変更一覧の保存先: %s" % out.resolve())
    else:
        findings, stats = judge(
            slides, cfg, prs.slide_width, prs.slide_height, slide_backgrounds(prs)
        )
        if structure_findings is None:
            structure_findings = judge_structure(
                slides, cfg, prs.slide_width, prs.slide_height
            )
        report = build_report(
            src,
            prs,
            slides,
            findings,
            stats,
            cfg,
            verdicts,
            candidates,
            structure_findings,
        )
        out = (
            Path(args.out)
            if args.out
            else Path.cwd() / (src.stem + "_reslide_report.md")
        )
        out.write_text(report, encoding="utf-8")
        print(report)
        print("レポート保存先: %s" % out.resolve())
    if sha256(src) != before:
        print("警告: 元ファイルのハッシュが変わっています", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
