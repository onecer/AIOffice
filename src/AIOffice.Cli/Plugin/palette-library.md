# aioffice palette library — pick a *usage contract*, not just colors

A palette is **proportions + roles**, not a swatch list. Apply every direction as:
- **bg** ≈ 60% of the surface — negative space is the look. Don't fill it.
- **panel** — cards / bands / table fills that sit on the bg.
- **ink** — primary text. **muted** — secondary text, captions, axis labels.
- **accent1..3** — used at **<8%** coverage: one accent bar, one highlighted bar in a chart, one big number. Never paint a slide in accents. No rainbow — 1 bg + 1 panel + 1 ink + 1 muted + ≤3 accents.

If the user gave brand colors / a logo / a website, **derive from those first** — a dark from the brand → bg, a tint of it → panel, near-black/white → ink, 1–3 brand colors → accents. The menu below is for when there is no brand.

| Direction | bg | panel | ink | muted | accents | temperament · best-for · analogy |
|---|---|---|---|---|---|---|
| Cool corporate | `FFFFFF` | `F4F8FD` | `0F2742` | `64748B` | `2563EB` `0EA5E9` `059669` | calm, trustworthy · B2B / finance / consulting · *like a McKinsey or Stripe deck* |
| Warm editorial | `FBF7EF` | `FFFFFF` | `1C1917` | `78716C` | `C2410C` `B45309` `1D4ED8` | human, considered · long-form / thought-leadership · *like an Economist feature spread* |
| Midnight tech | `0B1220` | `16213A` | `F8FAFC` | `94A3B8` | `38BDF8` `818CF8` `34D399` | modern, energetic · dev-tools / AI / launches · *like a Vercel or Linear launch* |
| Dark-cinematic | `0A0E27` | `1E293B` | `F1F5F9` | `94A3B8` | `14B8A6` `D4AF37` | premium, dramatic (one lit subject + gold) · keynote / film / luxury · *like an Apple product film* |
| Bold mono | `111111` | `1C1C1C` | `FFFFFF` | `A3A3A3` | `FACC15` (single) | stark, confident · manifestos / posters / single-idea · *like a brutalist or Off-White poster* |
| Soft product | `F8FAFC` | `FFFFFF` | `0F172A` | `64748B` | `6366F1` `EC4899` `14B8A6` | friendly, approachable · consumer SaaS / community · *like a Notion or Figma deck* |
| Jewel-tone | `FFFDF7` | `FFFFFF` | `1A1A1A` | `78716C` | `047857` `D4AF37` `7C3AED` | rich, luxurious · heritage / fashion / private banking · *like a Cartier or private-bank report* |
| Frost-ice | `F8FBFD` | `FFFFFF` | `0F2742` | `64748B` | `3B82F6` `0EA5E9` | clean, clinical · medical / health / SaaS · *like a Withings or clinical brief* |
| Nature-organic | `F7F4EC` | `FFFFFF` | `14310F` | `6B7280` | `166534` `D4AF37` `84CC16` | grounded, sustainable · wellness / food / ESG · *like a Patagonia field report* |
| Editorial-mono ink | `FBFBF9` | `FFFFFF` | `111111` | `6B7280` | `B91C1C` (single red) | austere, type-forward · academic / legal / press · *like a broadsheet front page* |

## Typography pairing (vary it to the tone)
Set the master `majorFont`/`minorFont` (pptx) or `/theme {majorFont,minorFont}` (docx):
- **neutral sans** (Inter / Helvetica) — safe default, any audience.
- **editorial** — a serif display (Georgia / Source Serif) + sans body.
- **modern geometric** — Space Grotesk / Sora / Sora display + sans body — tech, launches.
- **technical** — add a mono (JetBrains Mono / IBM Plex Mono) for labels, code, KPI numbers.
- **CJK** — pick a CJK family *explicitly* and pair it with a Latin family: e.g. `Source Han Sans` / `Noto Sans CJK` / `PingFang SC` / `Microsoft YaHei` for Chinese, plus a Latin family for numerals & loanwords. A Latin-only display font will fall back (ugly) on CJK glyphs. Embed `.ttf` fonts (`add /fonts type:font`) so the deck renders right off-machine.

## How to use a palette across a deck
Write the chosen row into `aioffice-spec.md` and **re-read it before every slide**. On pptx set it once on the master (`set /master[1] {background:"{bg}", accent1, accent2, accent3, majorFont, minorFont}`), then each slide inherits. The discipline that makes a deck look designed is *consistency under the proportion rule* — same bg everywhere, accents rationed, ink/muted never swapped.
