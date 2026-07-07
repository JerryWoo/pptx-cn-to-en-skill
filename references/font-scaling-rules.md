# Font Scaling Rules for Chinese → English PPTX Translation

## Why Font Scaling Is Needed

English text is typically 1.3–1.8× longer than the equivalent Chinese text in visual width,
because:
- CJK characters each occupy ~1 em of width
- Latin characters average ~0.6 em each (spaces count too)

This means a text box that fits "智慧公园管理平台" (8 CJK chars ≈ 8 em) may overflow when
replaced with "Smart Park Management Platform" (30 chars × 0.6 = 18 em).

## Algorithm (implemented in scripts/pptx_translate.py)

### Step 1 — Estimate width ratio per shape

For each `<p:sp>` (shape):
- Sum the `char_width()` of all translated paragraphs that changed
- Compare to the `char_width()` of the original text

```python
def char_width(text):
    w = 0.0
    for c in text:
        w += 1.8 if is_cjk(c) else 0.6
    return w

ratio = translated_width / original_width
```

### Step 2 — Decide whether to scale

| ratio | action |
|-------|--------|
| ≤ 1.15 | no change (within tolerance) |
| 1.15 – 1.92 | scale = OVERFLOW_THRESHOLD / ratio (proportional) |
| > 1.92 | scale = MIN_SCALE (0.60) — hard floor |

### Step 3 — Apply scale to `sz` attributes

All explicit `sz` (hundredths of a point, e.g. 2400 = 24pt) in `<a:rPr>` elements within
the shape are scaled:

```python
new_sz = max(int(original_sz * scale // 100 * 100), 600)  # minimum 6pt
```

Rounding to the nearest 100 units (1pt) avoids fractional point sizes that some renderers
handle poorly.

## Empirical Data (from a sample Chinese deck)

Observed manual adjustments showed these patterns:

| slide | original sz | adjusted sz | scale |
|-------|------------|------------|-------|
| 3 body bullets | 1200 | 1050 | 0.875 |
| 3 section labels | 1600 | 1200 | 0.750 |
| 4 vision bullets | 1600 | 1200 | 0.750 |
| 5 challenge text | 1200 | 1100 | 0.917 |
| 5 title | 3600 | 2600 | 0.722 |
| 9 district labels | 1200 | 800  | 0.667 |
| 10 data labels | 1400 | 900  | 0.643 |
| 10 section title | 2000 | 1100 | 0.550 |

The algorithm targets the observed average of ~0.75 for heavily-text shapes and ~0.90 for
lightly-text shapes, with the proportional formula producing per-shape tailored values.

## Tuning Parameters

Edit `scripts/pptx_translate.py` to adjust:

```python
OVERFLOW_THRESHOLD = 1.15   # expand tolerance to reduce unnecessary shrinking
MIN_SCALE = 0.60            # hard floor — raise if 6pt text is too small to read
```

For presentations with very dense Chinese (policy docs, legal text), consider lowering
`OVERFLOW_THRESHOLD` to 1.10 for more aggressive shrinking.
