# pptx-cn-to-en-skill

A reusable [WorkBuddy](https://www.codebuddy.cn/) skill that translates Chinese text inside a
PowerPoint (`.pptx`) file into English **while preserving the original layout**, and **automatically
shrinks font sizes** when the English text would overflow a text box.

It is designed for anyone who needs to turn a Chinese presentation (project proposals, product
introductions, reports, decks) into an English version without manually re-fitting every text box.

## Features

- **Full-slide translation** — covers slide body text, tables, SmartArt (`diagrams/data*.xml`),
  charts (`charts/chart*.xml`), and speaker notes (`notesSlides/*.xml`).
- **Layout preservation** — the PPTX is unpacked to XML, translated in place, and repacked, so
  colors, positions, shapes, and animations stay intact.
- **Auto font scaling** — English is typically 1.3–1.8× wider than Chinese. The script estimates
  the visual width ratio per shape and shrinks font sizes proportionally (down to a 60% floor)
  to prevent overflow. See `references/font-scaling-rules.md`.
- **Custom terminology** — supply a JSON dictionary for company names, product names, or
  domain jargon that must be translated precisely. User terms override the built-in dictionary.
- **Graceful fallback** — if the online translation backend is unavailable, the script still
  applies your terminology dictionary (no silent data loss).

## Requirements

- Python 3.10+
- The **`pptx` skill** (from the document-skills plugin) must be installed — it provides the
  `unpack.py` / `pack.py` / `clean.py` utilities used to read and write the PPTX container.
- Python packages: `lxml` (required), `requests` (optional, enables online translation).
- Network access to the [MyMemory](https://mymemory.translated.net/) translation API for
  automatic translation (optional but recommended).

## Installation

This is a WorkBuddy skill. Copy or symlink it into your skills directory:

```bash
# user-level skills (available across all projects)
cp -r pptx-cn-to-en-skill ~/.workbuddy/skills/

# or project-level skills (shared with a team on the same project)
cp -r pptx-cn-to-en-skill <your-project>/.workbuddy/skills/
```

Make sure the `pptx` skill is also installed so the script can find its `scripts/` directory.
If it lives somewhere non-standard, set the environment variable:

```bash
export PPTX_SKILL_SCRIPTS="/path/to/pptx/skill/scripts"
```

## Usage

### As a WorkBuddy skill

Just ask the agent in natural language, for example:

- "把这份中文方案.pptx 翻译成英文"
- "Translate this PowerPoint to English"
- "请将 D:\项目\中文方案.pptx 翻译成英文，专有名词：公司名称 → CompanyName"

The agent will locate the `.pptx`, optionally ask for custom terms, run the translation, and
present the resulting `_EN.pptx`.

### Directly from the command line

```bash
python scripts/pptx_translate.py input.pptx output.pptx
python scripts/pptx_translate.py input.pptx output.pptx --term-file terms.json
```

If `output.pptx` is omitted in spirit, the script writes `<input>_EN.pptx` next to the source.

## Custom terminology

Create a JSON file mapping source Chinese terms to their exact English translations:

```json
{
  "公司名称": "CompanyName",
  "产品名称": "ProductName",
  "核心算法": "Core Algorithm"
}
```

Pass it with `--term-file`. Longer matches are applied first, and your terms always win over the
built-in dictionary.

The built-in dictionary already covers ~50 common IT / business terms
(平台→Platform, 系统→System, 人工智能→Artificial Intelligence, 数字孪生→Digital Twin, …).

## How it works

1. Unpack the PPTX into its XML components (`pptx` skill's `unpack.py`).
2. For each slide and auxiliary XML file:
   - Merge the per-run text of every paragraph.
   - Apply custom + built-in terminology (longest match first).
   - Translate remaining Chinese via the MyMemory API (chunked, with timeout + cache).
   - Write the translated text back into the first run and clear the others.
3. For each shape with translated text, estimate the width ratio and shrink font sizes if needed.
4. Clean orphaned files and remove broken internal relationships (common in decks with missing
   media).
5. Repack into the output PPTX (`pptx` skill's `pack.py`).

## Tuning

Edit the constants at the top of `scripts/pptx_translate.py`:

```python
OVERFLOW_THRESHOLD = 1.15   # trigger shrinking when translated text is >15% wider
MIN_SCALE = 0.60            # never shrink below 60% of the original font size
```

For very dense Chinese text (legal, policy), lower `OVERFLOW_THRESHOLD` to ~1.10 for more
aggressive shrinking.

## Limitations

- Translations rely on a general-purpose machine translator; review domain-specific slides.
- This script translates text — it does not redesign layouts. Severely long English strings may
  still need a manual tweak after translation.
- Embedded images, videos, and OLE objects are left untouched (only their text labels are
  translated).

## License

Released under the [MIT License](./LICENSE).
