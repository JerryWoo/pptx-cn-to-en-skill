#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pptx_translate.py — PPTX 中文→英文翻译脚本（核心引擎）

用法：
    python pptx_translate.py <input.pptx> <output.pptx> [--term-file terms.json]

功能：
1. 解压 PPTX
2. 对每张幻灯片的所有段落：合并 run 文本 → 字典翻译 → 写回第一个 run
3. 对每个 shape：若翻译后文字长度显著增加，按比例缩小 sz（字号）
4. 打包输出

字号缩放逻辑（来自实测数据）：
    比例 = 原始中文长度 / 翻译后英文长度（近似字符当量）
    中文按 1.8 宽度当量计算（双字节），英文字符按 0.6 计算
    当超出判定阈值 > 1.15 时触发缩小，最低缩到原字号的 60%
"""

import argparse
import json
import re
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

from lxml import etree
try:
    import requests
except Exception:  # optional dependency
    requests = None

# ── 命名空间 ────────────────────────────────────────────────────────────────
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
A = lambda tag: f"{{{NS_A}}}{tag}"
DGM = lambda tag: f"{{{NS_DGM}}}{tag}"

# ── 内置翻译字典（通用 IT / 智慧城市术语） ─────────────────────────────────
BUILTIN_TERMS: dict[str, str] = {
    # 通用
    "简介": "Introduction",
    "背景": "Background",
    "目标": "Goals",
    "方案": "Solution",
    "功能": "Features",
    "系统": "System",
    "平台": "Platform",
    "数据": "Data",
    "管理": "Management",
    "服务": "Services",
    "安全": "Security",
    "智能": "Smart",
    "智慧": "Intelligent",
    "分析": "Analytics",
    "监控": "Monitoring",
    "告警": "Alert",
    "设备": "Device",
    "传感器": "Sensor",
    "摄像头": "Camera",
    "视频": "Video",
    "人工智能": "Artificial Intelligence",
    "大数据": "Big Data",
    "物联网": "Internet of Things",
    "数字孪生": "Digital Twin",
    "云计算": "Cloud Computing",
    "边缘计算": "Edge Computing",
    "区块链": "Blockchain",
    # 常见组合
    "解决方案": "Solution",
    "管理平台": "Management Platform",
    "运营中心": "Operations Center",
    "应用场景": "Application Scenarios",
    "核心优势": "Core Advantages",
    "技术架构": "Technical Architecture",
    "部署方案": "Deployment Solution",
    "总结": "Summary",
    "联系我们": "Contact Us",
}

_CN_RE = re.compile(r"[\u4e00-\u9fff\u3400-\u4dbf]")


def has_chinese(text: str) -> bool:
    return bool(_CN_RE.search(text))


def char_width(text: str) -> float:
    """估算文字视觉宽度当量（CJK≈1.8，ASCII≈0.6）。"""
    w = 0.0
    for c in text:
        w += 1.8 if "\u4e00" <= c <= "\u9fff" or "\u3400" <= c <= "\u4dbf" else 0.6
    return w


def build_translator(extra_terms: dict) -> tuple[dict, list]:
    """合并内置词典 + 用户词典，返回 (map, sorted_pairs)。"""
    merged = {**BUILTIN_TERMS, **extra_terms}
    sorted_pairs = sorted(merged.items(), key=lambda x: -len(x[0]))
    return merged, sorted_pairs


class TranslationEngine:
    def __init__(self, sorted_pairs: list, use_online: bool = True):
        self.sorted_pairs = sorted_pairs
        self.cache: dict[str, str] = {}
        self.use_online = use_online and requests is not None
        self.counter = 0

    def apply_terms(self, text: str) -> str:
        result = text
        for cn, en in self.sorted_pairs:
            if cn in result:
                result = result.replace(cn, en)
        return result

    def split_text(self, text: str, max_len: int = 420) -> list[str]:
        if len(text) <= max_len:
            return [text]
        parts = re.split(r"(?<=[。！？；;.!?])", text)
        chunks: list[str] = []
        current = ""
        for part in parts:
            if not part:
                continue
            if len(current) + len(part) > max_len and current:
                chunks.append(current)
                current = part
            else:
                current += part
        if current:
            chunks.append(current)
        # fallback for very long chunks without punctuation
        final: list[str] = []
        for chunk in chunks:
            if len(chunk) <= max_len:
                final.append(chunk)
            else:
                final.extend([chunk[i:i + max_len] for i in range(0, len(chunk), max_len)])
        return final

    def online_translate_once(self, text: str) -> str | None:
        if not self.use_online:
            return None
        try:
            resp = requests.get(
                "https://api.mymemory.translated.net/get",
                params={"q": text, "langpair": "zh-CN|en-US"},
                timeout=8,
            )
            resp.raise_for_status()
            data = resp.json()
            translated = data.get("responseData", {}).get("translatedText")
            if translated:
                return translated
        except Exception as exc:
            print(f"  WARN: online translation failed fast: {exc}", file=sys.stderr, flush=True)
        return None

    def translate(self, text: str) -> str:
        if not text or not has_chinese(text):
            return text
        if text in self.cache:
            return self.cache[text]

        protected = self.apply_terms(text)
        if not has_chinese(protected):
            self.cache[text] = protected
            return protected

        translated_parts = []
        all_ok = True
        for chunk in self.split_text(protected):
            if has_chinese(chunk):
                tr = self.online_translate_once(chunk)
                if tr:
                    translated_parts.append(tr)
                    self.counter += 1
                    if self.counter % 25 == 0:
                        print(f"  translated segments: {self.counter}", flush=True)
                    time.sleep(0.05)
                else:
                    all_ok = False
                    translated_parts.append(chunk)
            else:
                translated_parts.append(chunk)
        result = " ".join(part.strip() for part in translated_parts if part is not None).strip()
        if not result:
            result = protected
        self.cache[text] = result
        return result


def translate_str(text: str, engine: TranslationEngine) -> str:
    return engine.translate(text)


# ── 段落级翻译 ───────────────────────────────────────────────────────────────

def translate_paragraph(para, engine: TranslationEngine) -> tuple[bool, float, float]:
    """
    翻译段落内所有 run 的合并文本。
    返回 (changed, orig_width, translated_width)
    """
    runs = para.findall(A("r"))
    if not runs:
        return False, 0.0, 0.0

    run_data = []
    for r in runs:
        t = r.find(A("t"))
        run_data.append((r, t, t.text or "" if t is not None else ""))

    full_text = "".join(d[2] for d in run_data)
    if not has_chinese(full_text):
        return False, 0.0, 0.0

    translated = translate_str(full_text, engine)
    if translated == full_text:
        return False, 0.0, 0.0

    orig_w = char_width(full_text)
    new_w = char_width(translated)

    first_written = False
    for r, t, _ in run_data:
        if t is not None:
            if not first_written:
                t.text = translated
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                first_written = True
            else:
                t.text = ""
    return True, orig_w, new_w


# ── shape 级字号缩放 ─────────────────────────────────────────────────────────

OVERFLOW_THRESHOLD = 1.15  # 翻译后宽度超出原来 15% 时触发
MIN_SCALE = 0.60            # 最多缩到原字号 60%


def scale_shape_fonts(sp, scale: float) -> None:
    """对 shape 内所有显式 sz 属性按 scale 比例缩小（向下取整到 100 单位）。"""
    for rpr in sp.iter(A("rPr")):
        sz_str = rpr.get("sz")
        if sz_str:
            original = int(sz_str)
            new_sz = max(int(original * scale // 100 * 100), 600)  # 最小 6pt (600)
            if new_sz != original:
                rpr.set("sz", str(new_sz))


def process_shape_translation(sp, engine: TranslationEngine) -> None:
    """翻译 shape 内所有段落，并在必要时缩小字号。"""
    total_orig_w = 0.0
    total_new_w = 0.0

    for para in sp.iter(A("p")):
        changed, orig_w, new_w = translate_paragraph(para, engine)
        if changed:
            total_orig_w += orig_w
            total_new_w += new_w

    if total_orig_w > 0 and total_new_w > 0:
        ratio = total_new_w / total_orig_w
        if ratio > OVERFLOW_THRESHOLD:
            # 计算缩放系数：使翻译后内容大致恢复原宽度
            scale = max(MIN_SCALE, OVERFLOW_THRESHOLD / ratio)
            scale_shape_fonts(sp, scale)


# ── XML 文件处理 ─────────────────────────────────────────────────────────────

def translate_text_element(elem, engine: TranslationEngine) -> bool:
    """Translate a single text element such as <a:t> or <dgm:t>."""
    if elem.text and has_chinese(elem.text):
        translated = translate_str(elem.text, engine)
        if translated != elem.text:
            elem.text = translated
            elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            return True
    return False


def translate_generic_text(root, engine: TranslationEngine) -> bool:
    """Translate text in non-shape XML regions (tables, SmartArt data, notes, charts)."""
    changed = False
    for tag in (A("t"), DGM("t")):
        for elem in root.iter(tag):
            if translate_text_element(elem, engine):
                changed = True
    return changed


def process_xml_file(xml_path: Path, engine: TranslationEngine, adjust_fonts: bool = True) -> bool:
    raw = xml_path.read_bytes()
    try:
        root = etree.fromstring(raw)
    except Exception as e:
        print(f"  PARSE ERROR ({xml_path.name}): {e}", file=sys.stderr)
        return False

    changed = False

    # Prefer shape-level translation for slides so font scaling can use per-shape width ratios.
    if adjust_fonts:
        for sp in root.iter(A("sp")):
            pre_state = etree.tostring(sp)
            process_shape_translation(sp, engine)
            if etree.tostring(sp) != pre_state:
                changed = True

    # Catch all remaining a:t/dgm:t text, including tables, charts, SmartArt data, notes.
    if translate_generic_text(root, engine):
        changed = True

    if changed:
        xml_path.write_bytes(
            etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        )
    return changed


def source_xml_for_rels_file(rel_file: Path) -> Path:
    """Map ppt/slides/_rels/slide1.xml.rels to ppt/slides/slide1.xml."""
    return rel_file.parent.parent / rel_file.name[:-5]


def remove_rid_references(source_xml: Path, bad_rids: set[str]) -> int:
    """Remove attributes in source XML that reference missing relationship IDs."""
    if not source_xml.exists() or not bad_rids:
        return 0
    try:
        root = etree.fromstring(source_xml.read_bytes())
    except Exception:
        return 0
    removed = 0
    for elem in root.iter():
        for attr in list(elem.attrib):
            if attr.endswith("}id") and elem.attrib.get(attr) in bad_rids:
                del elem.attrib[attr]
                removed += 1
    if removed:
        source_xml.write_bytes(
            etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        )
    return removed


def clean_broken_relationships(work_dir: Path) -> int:
    """Remove relationship entries whose internal Target file does not exist and clear stale r:id refs."""
    removed = 0
    rel_files = list(work_dir.glob("**/_rels/*.rels"))
    for rel_file in rel_files:
        # Never touch package-level root relationships such as _rels/.rels;
        # removing those makes docProps and ppt/presentation.xml appear unreferenced.
        if rel_file.relative_to(work_dir).as_posix() == "_rels/.rels":
            continue
        try:
            root = etree.fromstring(rel_file.read_bytes())
        except Exception:
            continue
        changed = False
        source_xml = source_xml_for_rels_file(rel_file)
        source_dir = source_xml.parent
        bad_rids: set[str] = set()
        for rel in list(root):
            target = rel.get("Target")
            if not target or target.startswith(("http://", "https://", "mailto:")):
                continue
            target_path = (source_dir / target).resolve()
            if not target_path.exists():
                rid = rel.get("Id")
                if rid:
                    bad_rids.add(rid)
                root.remove(rel)
                removed += 1
                changed = True
        if changed:
            remove_rid_references(source_xml, bad_rids)
            rel_file.write_bytes(
                etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            )
    return removed


# ── 主流程 ───────────────────────────────────────────────────────────────────

def find_pptx_skill():
    """找到 pptx skill 的 scripts 目录（用于调用 unpack/pack/clean）。

    会在常见安装位置中查找，也可通过环境变量 PPTX_SKILL_SCRIPTS 显式指定。
    """
    import os
    env_path = os.environ.get("PPTX_SKILL_SCRIPTS")
    if env_path and (Path(env_path) / "office" / "unpack.py").exists():
        return Path(env_path)

    base = Path.home()
    candidates = [
        base / ".workbuddy/plugins/marketplaces/cb_teams_marketplace/plugins/document-skills/skills/pptx/scripts",
        base / ".workbuddy/skills/pptx/scripts",
        base / ".workbuddy/plugins/marketplaces/document-skills/skills/pptx/scripts",
    ]
    for c in candidates:
        if (c / "office" / "unpack.py").exists():
            return c

    # 退而求其次：在用户目录下做广域搜索
    matches = list(base.glob("**/document-skills/skills/pptx/scripts/office/unpack.py"))
    if matches:
        return matches[0].parent

    raise FileNotFoundError(
        "Cannot find pptx skill scripts. Please install the pptx skill, "
        "or set the PPTX_SKILL_SCRIPTS environment variable to its scripts/ directory."
    )


def run_pptx_script(scripts_dir: Path, script: str, *args: str) -> None:
    cmd = [sys.executable, str(scripts_dir / script)] + list(args)
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.stdout:
        print(result.stdout.rstrip())
    if result.returncode != 0:
        print(result.stderr.rstrip(), file=sys.stderr)
        raise RuntimeError(f"{script} failed with exit code {result.returncode}")


def translate_pptx(input_path: str, output_path: str, term_file: str | None = None) -> None:
    input_path = Path(input_path)
    output_path = Path(output_path)

    # 加载用户词典
    extra_terms: dict = {}
    if term_file:
        with open(term_file, encoding="utf-8") as f:
            extra_terms = json.load(f)
    _, sorted_pairs = build_translator(extra_terms)
    engine = TranslationEngine(sorted_pairs, use_online=True)

    scripts_dir = find_pptx_skill()

    with tempfile.TemporaryDirectory() as tmp:
        work_dir = Path(tmp) / "unpacked"

        # 1. 解压
        print(f"Unpacking {input_path.name} …")
        run_pptx_script(scripts_dir, "office/unpack.py", str(input_path), str(work_dir))

        # 2. 翻译 & 字号调整
        slides_dir = work_dir / "ppt" / "slides"
        slide_files = sorted(
            [f for f in slides_dir.glob("slide*.xml") if "Layout" not in f.name],
            key=lambda p: int(re.search(r"\d+", p.stem).group()),
        )

        print(f"Translating {len(slide_files)} slides …")
        changed_count = 0
        for sf in slide_files:
            ok = process_xml_file(sf, engine, adjust_fonts=True)
            if ok:
                changed_count += 1
                print(f"  {sf.name}: translated + font-adjusted")

        # Translate SmartArt/diagram data and other non-slide XML text.
        extra_xml_files = []
        for pattern in (
            "ppt/diagrams/data*.xml",
            "ppt/charts/chart*.xml",
            "ppt/notesSlides/notesSlide*.xml",
        ):
            extra_xml_files.extend(work_dir.glob(pattern))
        extra_changed = 0
        for xf in sorted(extra_xml_files):
            if process_xml_file(xf, engine, adjust_fonts=False):
                extra_changed += 1

        print(f"  {changed_count}/{len(slide_files)} slides updated; {extra_changed} auxiliary XML files updated.")

        # 3. 清理
        run_pptx_script(scripts_dir, "clean.py", str(work_dir))
        broken = clean_broken_relationships(work_dir)
        if broken:
            print(f"Removed {broken} broken internal relationship(s).")

        # 4. 打包
        print(f"Packing → {output_path.name} …")
        run_pptx_script(
            scripts_dir,
            "office/pack.py",
            str(work_dir),
            str(output_path),
            "--original",
            str(input_path),
        )

    print(f"Done! Output: {output_path}")


# ── CLI ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Translate Chinese text in a PPTX to English, with automatic font-size adjustment."
    )
    parser.add_argument("input", help="Source PPTX (Chinese)")
    parser.add_argument("output", help="Output PPTX (English)")
    parser.add_argument(
        "--term-file",
        help="JSON file mapping Chinese terms to English (overrides / extends built-in dictionary)",
        default=None,
    )
    args = parser.parse_args()
    translate_pptx(args.input, args.output, args.term_file)
