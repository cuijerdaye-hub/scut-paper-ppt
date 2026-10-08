# 华工蓝模板画像

Use this reference when generating or editing decks with `scut-paper-ppt`.

## File Structure

- Bundled template: `assets/scut-blue-template.pptx`
- Slides: 19 sample slides
- Slide masters: 1
- Slide layouts: 13
- Themes: 3
- Canvas: 13.33 x 7.5 inches, 16:9

The visible template design is mostly built on manually arranged sample slides, images, shapes, and text boxes. The master/layout placeholders are generic Office placeholders and are not enough to recreate the visual style by themselves.

## Visual System

- Primary blue: `#2D4F8F`
- Secondary blue: `#76ABDC`
- Accent yellow: `#FFCC00`
- Supporting colors: `#FFFFFF`, `#000000`
- Preferred fonts: `思源黑体 CN`, `思源黑体 Heavy`
- Fallback Chinese font: `微软雅黑`
- Common sizes: 60-72 pt section numerals, 44-48 pt title text, 20-24 pt section headers, 12-16 pt body text.

## Sample Slide Mapping

| Purpose | Template slide(s) | Notes |
|---|---:|---|
| Cover | 1, 7 | Slide 7 includes speaker and date fields; use it as default cover. |
| Directory | 8, 9, 10 | Slide 8 is the default four-part contents page. Slide 10 can be reused as a roadmap page. |
| Section divider | 11, 12, 13, 14 | Large numeral, title, subtitle, and short summary. |
| Overview / contribution | 15 | Main paragraph plus three numbered cards. |
| Multi-module content | 16 | Four horizontal modules; good for methods, framework, or findings. |
| Three-card content | 17 | Three visual cards; good for datasets, experiments, or result groups. |
| Image + summary | 18 | Large image area plus headline, bullets, and long note. |
| Closing | 19 | Closing thanks and speaker line. |

## Default Deck Narrative

Default generated structure keeps six complete paper-report parts:

1. Cover
2. Contents with six parts
3. Section 01: Research background
4. Background overview
5. Section 02: Problem and contributions
6. Problem / contribution content
7. Section 03: Method framework
8. Method framework content
9. Section 04: Experimental design
10. Experiment / evidence content
11. Section 05: Results and discussion
12. Results content
13. Section 06: Conclusion and outlook
14. Closing

The template only has four section-divider sample slides, so the generator clones those samples when six dedicated section dividers are required.

## Placeholder Cleanup

Generated decks must not retain these sample strings:

- `鲤工仔`
- `Your Title Here SCUT`
- `XX年XX月`
- `华工蓝幻灯片模板`

It is acceptable to retain official school identity text such as `South China University Of Technology`.

## Quality Gates

Run both validators after generation and after any manual or COM-based edit:

```bash
python scripts/validate_scut_deck.py output.pptx
python scripts/pptx_deck_qa.py output.pptx --expected-slides 14 --report output.qa.json --strict
python scripts/preview_deck_qa.py output-preview --expected-slides 14 --dark-slide 3 --evidence-slide 10 --report output.preview-qa.json --strict
python scripts/vision_review_packet.py output-preview --output-dir output-vision-review --pptx-qa output.qa.json --preview-qa output.preview-qa.json --strict
```

A valid first pass should satisfy:

- Canvas remains `13.33 x 7.5` inches.
- Deck has at least 14 slides.
- Contents lists all six chapter titles.
- Six section divider pages exist and are numbered `01` through `06`.
- No sample placeholders remain.
- PPTX package status is `pass`: no broken relationships, orphan media, unresolved placeholders, `layout:` markers, `????` mojibake, or fully off-canvas shapes.
- Preview-image status is `pass`: no missing/low-resolution/inconsistent previews, no nearly blank slides, no oversized white area on declared dark pages, and no missing SCUT-blue accent on declared evidence pages.
- Vision review packet status is `pass` and includes `vision-contact-sheet.png`, `vision-review-packet.json`, and `vision-review-prompt.md`.

When visual QA is possible, export the PPTX to PDF/PNG, build the Vision review packet, and inspect the cover, contents, one dark section divider, one light section divider, one content slide, one evidence/screenshot slide, and the closing page. Look for wrong master use, text overlap, screenshot legibility, source caption placement, and highlight boxes that obscure text.

For shorter custom reports, adjust `--expected-slides` to the actual planned count and use `--require-source` for every required evidence caption.
