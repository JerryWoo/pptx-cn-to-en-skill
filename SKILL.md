---
name: pptx-cn-to-en-translator
description: Translate Chinese text in PowerPoint presentations to English while preserving layout and automatically adjusting font sizes to prevent text overflow. Use when the user asks to translate a Chinese PPTX to English, or mentions terms like "PPTX翻译", "PPT中译英", "中文方案英文版". Triggers include file paths ending in .pptx with translation requests.
agent_created: true
---

# PPTX Chinese → English Translator

Translate all Chinese text in a PowerPoint presentation to English while:
1. Preserving the original layout, colors, and formatting
2. Automatically shrinking font sizes when English text would overflow the text boxes
3. Supporting custom terminology dictionaries for domain-specific translations

## When to Use This Skill

Invoke this skill when:
- The user asks to translate a Chinese PPTX file to English
- The user provides a .pptx file path and mentions translation ("翻译", "英文版", "translate to English")
- The user requests bilingual slides or English-only output from Chinese source material

**Explicit triggers:**
- "请将这个PPT翻译成英文"
- "Translate this PowerPoint to English"
- "把这份中文方案做成英文版"
- "Help me create an English version of this presentation"

## Core Workflow

### Step 1: Understand Requirements

Ask the user:
1. **Source file path** — full path to the Chinese .pptx
2. **Output file path** — where to save the translated .pptx (default: `<original>_EN.pptx` in same directory)
3. **Custom terminology** — any domain-specific terms that need precise translation (optional)

If the user provides custom terms, create a temporary JSON dictionary file:

```json
{
  "公司名称": "CompanyName",
  "产品名称": "ProductName"
}
```

### Step 2: Run Translation Script

Execute the core translation engine:

```bash
python scripts/pptx_translate.py <input.pptx> <output.pptx> [--term-file terms.json]
```

**What the script does:**
1. Unpacks the PPTX into XML components (via pptx skill's unpack.py)
2. For each slide and auxiliary XML file:
   - Merges run-level text in each paragraph
   - Applies custom terminology first (longest-match-first)
   - Sends remaining Chinese text to MyMemory translation API with chunking and timeout
   - Writes translated text back to first run, clears others
   - Also translates SmartArt (`ppt/diagrams/data*.xml`), charts (`ppt/charts/chart*.xml`), and notes (`ppt/notesSlides/notesSlide*.xml`)
3. For each shape with translated text:
   - Calculates visual width ratio (CJK chars ≈ 1.8em, Latin ≈ 0.6em each)
   - If translated width > 1.15× original, shrinks font size proportionally
   - Minimum scale: 0.60 (no smaller than 60% of original size)
4. Cleans orphaned files (via pptx skill's clean.py)
5. Removes broken internal relationships and stale `r:id` references (common in PPTX files with missing videos/media)
6. Repacks into output.pptx (via pptx skill's pack.py)

**Script dependencies:**
- Requires the `pptx` skill to be installed (provides unpack/pack/clean utilities)
- Uses lxml for safe XML parsing
- Uses MyMemory HTTP translation API as the online translation backend (with 8s per-request timeout, sentence chunking, and cache)
- Falls back to term-only replacement when online translation fails

### Step 3: Quality Check

After translation completes:

1. **Text extraction QA** — verify no Chinese characters remain in slide content:
   ```bash
   python -m markitdown output.pptx | grep -P '[\u4e00-\u9fff]' | grep -v 'Notes:'
   ```

   If Chinese text is found (excluding Notes section), review whether it's:
   - Missing from dictionary → add to custom terms and re-run
   - Intentionally preserved (brand names, proper nouns)
   - Edge case requiring manual fix

2. **Visual QA** — use subagent to inspect slide images:
   ```bash
   # Convert to PDF then images
   python scripts/office/soffice.py --headless --convert-to pdf output.pptx
   pdftoppm -jpeg -r 150 output.pdf slide
   ```

   Ask subagent to check for:
   - Text overflow (words cut off at text box edges)
   - Overlapping text
   - Excessive whitespace (font shrunk too aggressively)
   - Alignment or formatting issues

3. **Present to user** — use `present_files` tool to show the translated PPTX.

## Font Size Adjustment Logic

English text is typically 1.3–1.8× wider than Chinese text due to character density differences.
The script automatically detects overflow and shrinks fonts to fit.

**Algorithm:**
- Estimate visual width: CJK = 1.8 em/char, Latin = 0.6 em/char
- Calculate ratio: `translated_width / original_width`
- If ratio > 1.15 (overflow threshold):
  - Scale = min(1.15 / ratio, 0.60)  ← proportional shrink, hard floor at 60%
  - Apply to all `sz` attributes in the shape

**Tuning parameters** (edit `scripts/pptx_translate.py` if needed):
```python
OVERFLOW_THRESHOLD = 1.15   # trigger shrinking at 15% overflow
MIN_SCALE = 0.60            # never shrink below 60% of original size
```

See `references/font-scaling-rules.md` for empirical data and detailed algorithm explanation.

## Built-in Translation Dictionary

The script includes a general-purpose IT / common-business terminology dictionary covering ~50 common terms:
- 智慧/智能 → Smart/Intelligent
- 平台/系统/服务 → Platform/System/Services
- 人工智能/大数据/物联网 → AI/Big Data/IoT
- 管理/监控/告警 → Management/Monitoring/Alert

For domain-specific projects, **always ask the user** if there are custom terms (company names,
product names, technical jargon) that need precise translation.

## Custom Terminology Workflow

When the user provides custom terms:

1. **Create term file** — write a JSON dictionary:
   ```json
   {
     "source_term_cn": "Target Translation",
     "another_term": "Another Translation"
   }
   ```

2. **Merge behavior** — user terms override built-in dictionary (longest match wins)

3. **Pass to script** — `--term-file <path.json>` argument

**Example:**
```bash
# User request: "公司名称 → CompanyName, 产品名称 → ProductName"
cat > /tmp/custom_terms.json << 'EOF'
{
  "公司名称": "CompanyName",
  "产品名称": "ProductName"
}
EOF

python scripts/pptx_translate.py input.pptx output.pptx --term-file /tmp/custom_terms.json
```

## Troubleshooting

### Issue: Text still overflows after translation

**Cause:** Font shrinking hit the 60% floor but text is still too long.

**Fix:**
1. Check if translation can be shortened (use more concise English)
2. Lower `MIN_SCALE` to 0.50 in the script (risky — may become unreadable)
3. Manually adjust problematic slides in PowerPoint after translation

### Issue: Some Chinese text not translated

**Cause:** Missing from dictionary (neither built-in nor custom terms).

**Fix:**
1. Extract remaining Chinese: `python -m markitdown output.pptx | grep -P '[\u4e00-\u9fff]'`
2. Add to custom term file
3. Re-run translation

### Issue: Script fails with "Cannot find pptx skill"

**Cause:** The `pptx` skill (marketplace) is not installed.

**Fix:**
Install the `pptx` skill from the document-skills plugin, or place its `scripts/` directory
where `find_pptx_skill()` can discover it (see `scripts/pptx_translate.py`).

## Example Session

**User:** "请将 D:\项目\中文方案.pptx 翻译成英文，专有名词：公司名称 → CompanyName"

**Agent response:**
1. Create custom term file with `{"公司名称": "CompanyName"}`
2. Run: `python scripts/pptx_translate.py "D:\项目\中文方案.pptx" "D:\项目\中文方案_EN.pptx" --term-file /tmp/terms.json`
3. QA: Extract text and check for remaining Chinese
4. Present: `present_files(["D:\项目\中文方案_EN.pptx"])`

**Output:** Fully translated English PPTX with automatic font adjustments, ready for use.
