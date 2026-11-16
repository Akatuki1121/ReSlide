#!/usr/bin/env python3
"""ReSlide: .pptx を走査し、見やすさの問題をレポートする（フェーズ1）。

使い方: python reslide.py <対象.pptx> [--config 設定.json] [--out レポート.md]
       python reslide.py <対象.pptx> --fix   … 自動調整（フェーズ2）
       python reslide.py <対象.pptx> --split … 分割候補の判定だけをレポートする（フェーズ3）
       python reslide.py <対象.pptx> --fix --split … 自動調整に加えてページ分割も行う
       python reslide.py <対象.pptx> --reorder … 並び替え候補だけをレポートする（フェーズ4）
       python reslide.py <対象.pptx> --reorder-apply <候補番号> … 選んだ候補だけ並び替えて保存する
元の .pptx は一切変更しない（開くだけで保存はしない。実行前後のハッシュも照合する）。
--fix の結果は <元名>_reslide.pptx へ新規保存し、変更一覧を Markdown で出す。
"""

import argparse
import copy
import hashlib
import json
import math
import re
import sys
import unicodedata
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER

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
        self.paragraphs = (
            paragraphs  # [{"text","runs":[(n,size,src)],"run_els":[(要素,文字列)]}]
        )
        self.scale = scale
        self.inner_w = inner_w  # 枠の内側（余白を差し引いた）幅・EMU・絶対座標
        self.inner_h = inner_h  # 同・高さ

    @property
    def text(self):
        return "\n".join(p["text"] for p in self.paragraphs if p["text"].strip())


def read_text_unit(resolver, shape, txbody, label="", inner_w=0, inner_h=0):
    scale = autofit_scale(shape) if label == "" else 1.0
    paras = []
    for p_el in txbody.findall("a:p", NS):
        runs, run_els, text_parts = [], [], []
        for child in p_el:
            if child.tag in (A + "r", A + "fld"):
                t = child.find("a:t", NS)
                txt = (t.text or "") if t is not None else ""
                if txt:
                    size, src = resolver.resolve(shape, p_el, child)
                    runs.append((len(txt.replace("\n", "")), size, src))
                    run_els.append((child, txt))
                text_parts.append(txt)
            elif child.tag == A + "br":
                text_parts.append("\n")
        paras.append({"text": "".join(text_parts), "runs": runs, "run_els": run_els})
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
                        el.units.append(
                            read_text_unit(
                                resolver,
                                sh,
                                tb,
                                "セル(%d行%d列)" % (ri + 1, ci + 1),
                                iw,
                                ih,
                            )
                        )
        elif sh.has_text_frame:
            il, ir, it, ib = _insets(sh.text_frame)
            iw = (w - il - ir) * tf.sx
            ih = (h - it - ib) * tf.sy
            el.units.append(
                read_text_unit(
                    resolver, sh, sh._element.find("p:txBody", NS), "", iw, ih
                )
            )
        out.append(el)


def read_presentation(prs, resolver):
    slides = []
    for i, slide in enumerate(prs.slides, 1):
        elems = []
        walk(resolver, i, slide.shapes, Tf(), elems)
        slides.append(elems)
    return slides


# ---------------------------------------------------------------- 判定
def classify(elem, unit, cfg):
    """(役割, 根拠) を返す。役割: title / note / label / body"""
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
    # ラベル: 表のセルでなく、1段落のみ・改行なし・短いテキスト
    if unit.label == "":
        texts = [p["text"] for p in unit.paragraphs if p["text"].strip()]
        if len(texts) == 1 and "\n" not in texts[0]:
            n = len(texts[0].strip())
            if n <= cfg.get("label_max_chars", 30):
                return "label", "1段落・改行なし・%d文字のためラベル" % n
    return "body", "上記に該当しないため本文"


ROLE_JP = {"body": "本文", "label": "ラベル"}


def judge(slides, cfg, slide_w, slide_h):
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

    for sno, elems in enumerate(slides, 1):
        for e in elems:
            for u in e.units:
                if not u.text.strip():
                    continue
                stats["units"] += 1
                role, why = classify(e, u, cfg)
                floor_u = floors.get(role)
                small_chars = total = 0
                sizes = []
                for p in u.paragraphs:
                    for n, size, src in p["runs"]:
                        total += n
                        stats["chars_total"] += n
                        if size is None:
                            stats["chars_unresolved"] += n
                            stats["unresolved"].append(
                                (sno, e.name, u.label, p["text"][:20])
                            )
                            continue
                        eff = size * u.scale
                        stats["by_src"][src] = stats["by_src"].get(src, 0) + n
                        sizes.append(eff)
                        if floor_u is not None and eff < floor_u:
                            small_chars += n
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
                            "detail": "最小 %.1fpt / %sの下限 %gpt（%d/%d文字が下限未満）%s。役割=%s（%s）"
                            % (
                                min(sizes),
                                ROLE_JP[role],
                                floor_u,
                                small_chars,
                                total,
                                scale_note,
                                ROLE_JP[role],
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
    return findings, stats


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
    sizes = [
        size * unit.scale
        for p in unit.paragraphs
        for (n, size, src) in p["runs"]
        if size is not None
    ]
    if not sizes:
        return None
    cur_min = min(sizes)
    if cur_min >= floor:
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
    for sno, elems in enumerate(slides, 1):
        # 1. サイズの拡大（note と title は触らない）
        for e in elems:
            for u in e.units:
                if not u.text.strip():
                    continue
                role, why = classify(e, u, cfg)
                floor = floors.get(role)
                if floor is None:
                    continue
                target = _fit_size(u, floor)
                if target is None:
                    continue
                old_min = min(
                    size * u.scale
                    for p in u.paragraphs
                    for (n, size, s2) in p["runs"]
                    if size is not None
                )
                _apply_size(u, target)
                rec(
                    sno,
                    "サイズ",
                    e,
                    u.label,
                    "最小 %.1fpt → %.1fpt（役割=%s、下限 %gpt。枠に収まる範囲で拡大）"
                    % (old_min, target, ROLE_JP[role], floor),
                    head=u.text,
                )

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


def build_changes_report(path, out_name, changes):
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
    if not changes:
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


def _is_note_only(e, cfg):
    """テキストの役割がすべて注記（出典・ページ番号など）の要素か。

    注記はスライドの付属物として元のスライドに残す
    （フェーズ2で注記のサイズを触らないのと同じ扱い）。
    """
    seen = False
    for u in e.units:
        if not u.text.strip():
            continue
        seen = True
        role, _why = classify(e, u, cfg)
        if role != "note":
            return False
    return seen


def _split_movable(elems, heading, cfg):
    """分割で後半へ移す要素。読み順（上から、左から）に並べて返す。

    対象はスライド直下の要素だけ。グループ内の要素は親のグループごとで
    なければ移せないため対象にしない。画像は元のスライドに残し、
    表・グラフ・SmartArt は内部を割らず要素ごと移す。
    """
    out = []
    for e in elems:
        if e is heading or e.shape is None or e.in_group:
            continue
        if not (e.has_text or e.kind in SPLIT_MOVABLE_KINDS):
            continue
        if _is_note_only(e, cfg):
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
    total = 0
    for e in elems:
        if e is heading or _is_note_only(e, cfg):
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
    lines = []
    for e in elems:
        if e is heading or _is_note_only(e, cfg):
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
    L.append("")
    L.append("## フォントサイズの解決状況")
    L.append("")
    L.append("- 文字数（run単位で数えた全テキスト）: %d" % stats["chars_total"])
    L.append("- **未解決の文字数: %d**" % stats["chars_unresolved"])
    for src, n in sorted(stats["by_src"].items(), key=lambda x: -x[1]):
        mark = (
            "（既定値。明示指定が無く継承で決まった分）"
            if src in DEFAULT_SOURCES
            else ""
        )
        L.append("  - 解決元 %s: %d文字%s" % (src, n, mark))
    if stats["unresolved"]:
        L.append("")
        L.append(
            "未解決の箇所（サイズが取れなかったため判定対象外。黙って落とさず列挙）:"
        )
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
            L.append(
                "- 実行するには --reorder-apply <候補番号> を指定する。選ばない限り、ファイルは作られない"
            )
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
        "- 継承（レイアウト・マスター）で決まる値は、プレースホルダの対応付けが取れた範囲で解決する。表スタイル由来のサイズは見ない。"
    )
    L.append(
        "- 重なりは、テキストを持つ要素と、図・表・グラフ同士の組み合わせのみ判定（背景用の文字なし図形と線、図(画像)同士は除外）。"
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
    args = ap.parse_args(argv)
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8")
        except Exception:
            pass
    if args.reorder_apply is not None and args.split:
        print(
            "--reorder-apply と --split は併用できません。分割でスライドが増えると、候補番号が指す並びが変わるためです",
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
    if args.fix or chosen is not None:
        changes = []
        if args.fix:
            slide_objs = list(prs.slides)
            changes = apply_fixes(slides, cfg, prs.slide_width, prs.slide_height)
            if args.split:
                changes += apply_splits(prs, slide_objs, slides, verdicts, cfg)
        if chosen is not None:
            # 並び替えは最後に行う。--fix の調整はスライドの枚数も順序も変えないため、
            # 候補が指す番号は sldIdLst の並びとそのまま対応する
            changes.append(apply_reorder(prs, chosen))
        out_pptx = Path.cwd() / (src.stem + "_reslide.pptx")
        prs.save(str(out_pptx))
        report = build_changes_report(src, out_pptx.name, changes)
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
        findings, stats = judge(slides, cfg, prs.slide_width, prs.slide_height)
        report = build_report(
            src, prs, slides, findings, stats, cfg, verdicts, candidates
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
