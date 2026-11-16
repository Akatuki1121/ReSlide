#!/usr/bin/env python3
"""ReSlide フェーズ1: .pptx を読み取り専用で走査し、見やすさの問題をレポートする。

使い方: python reslide.py <対象.pptx> [--config 設定.json] [--out レポート.md]
元の .pptx は一切変更しない（開くだけで保存はしない。実行前後のハッシュも照合する）。
"""

import argparse
import hashlib
import json
import re
import sys
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

    def __init__(self, label, paragraphs, scale):
        self.label = label  # 例: "" / "セル(2,3)"
        self.paragraphs = paragraphs  # [{"text","chars":[(n,size,src)]}]
        self.scale = scale

    @property
    def text(self):
        return "\n".join(p["text"] for p in self.paragraphs if p["text"].strip())


def read_text_unit(resolver, shape, txbody, label=""):
    scale = autofit_scale(shape) if label == "" else 1.0
    paras = []
    for p_el in txbody.findall("a:p", NS):
        runs, text_parts = [], []
        for child in p_el:
            if child.tag in (A + "r", A + "fld"):
                t = child.find("a:t", NS)
                txt = (t.text or "") if t is not None else ""
                if txt:
                    size, src = resolver.resolve(shape, p_el, child)
                    runs.append((len(txt.replace("\n", "")), size, src))
                text_parts.append(txt)
            elif child.tag == A + "br":
                text_parts.append("\n")
        paras.append({"text": "".join(text_parts), "runs": runs})
    return TextUnit(label, paras, scale)


class Elem:
    def __init__(self, slide_no, kind, name, x, y, w, h):
        self.slide_no, self.kind, self.name = slide_no, kind, name
        self.x, self.y, self.w, self.h = x, y, w, h  # EMU・スライド絶対座標
        self.units = []
        self.ph_type = None  # PP_PLACEHOLDER 名
        self.in_group = False

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
                        el.units.append(
                            read_text_unit(
                                resolver, sh, tb, "セル(%d行%d列)" % (ri + 1, ci + 1)
                            )
                        )
        elif sh.has_text_frame:
            el.units.append(
                read_text_unit(resolver, sh, sh._element.find("p:txBody", NS))
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
    """(役割, 根拠) を返す。役割: title / note / body"""
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
    return "body", "上記に該当しないため本文"


def judge(slides, cfg, slide_w, slide_h):
    findings = {i: [] for i in range(1, len(slides) + 1)}
    stats = {
        "chars_total": 0,
        "chars_unresolved": 0,
        "unresolved": [],
        "by_src": {},
        "notes": [],
        "units": 0,
    }
    floor = cfg["body_min_pt"]
    tol = cfg["out_of_bounds_tolerance_mm"] * EMU_PER_MM

    for sno, elems in enumerate(slides, 1):
        for e in elems:
            for u in e.units:
                if not u.text.strip():
                    continue
                stats["units"] += 1
                role, why = classify(e, u, cfg)
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
                        if role == "body" and eff < floor:
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
                if role == "body" and small_chars:
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
                            "detail": "最小 %.1fpt / 本文の下限 %gpt（%d/%d文字が下限未満）%s。役割=本文（%s）"
                            % (min(sizes), floor, small_chars, total, scale_note, why),
                        }
                    )

        # はみ出し
        for e in elems:
            if e.kind == "線・コネクタ":
                continue
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


# ---------------------------------------------------------------- レポート
def build_report(path, prs, slides, findings, stats, cfg):
    L = []
    L.append("# ReSlide レポート: %s" % Path(path).name)
    L.append("")
    L.append(
        "- スライド数: %d / サイズ: %.0f x %.0f mm（単位はすべて mm、フォントは pt）"
        % (len(slides), mm(prs.slide_width), mm(prs.slide_height))
    )
    L.append("- 基準値: 本文の下限 %gpt（設定ファイルの値）" % cfg["body_min_pt"])
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
    L.append("## 対象外・前提")
    L.append("")
    L.append("- グラフ・SmartArt の内部の文字は読まない（要素の位置と大きさのみ）。")
    L.append("- 回転した要素は回転前の矩形で判定する。")
    L.append(
        "- 継承（レイアウト・マスター）で決まる値は、プレースホルダの対応付けが取れた範囲で解決する。表スタイル由来のサイズは見ない。"
    )
    L.append(
        "- 重なりは、テキストを持つ要素と、図・表・グラフ同士の組み合わせのみ判定（背景用の文字なし図形と線は除外）。"
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
        description="pptx の見やすさ問題をレポートする（読み取り専用）"
    )
    ap.add_argument("pptx")
    ap.add_argument(
        "--config", default=str(Path(__file__).with_name("reslide_config.json"))
    )
    ap.add_argument(
        "--out",
        default=None,
        help="Markdownレポートの出力先（既定: カレントに <名前>_reslide_report.md）",
    )
    args = ap.parse_args(argv)
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8")
        except Exception:
            pass
    src = Path(args.pptx)
    cfg = json.loads(Path(args.config).read_text(encoding="utf-8"))
    before = sha256(src)
    prs = Presentation(str(src))
    resolver = Resolver(prs)
    slides = read_presentation(prs, resolver)
    findings, stats = judge(slides, cfg, prs.slide_width, prs.slide_height)
    report = build_report(src, prs, slides, findings, stats, cfg)
    out = Path(args.out) if args.out else Path.cwd() / (src.stem + "_reslide_report.md")
    out.write_text(report, encoding="utf-8")
    print(report)
    print("レポート保存先: %s" % out.resolve())
    if sha256(src) != before:
        print("警告: 元ファイルのハッシュが変わっています", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
