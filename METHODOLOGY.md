# Methodology

This document explains exactly how the numbers in this project are produced, so
you can judge the comparison for yourself, criticise it, and reproduce it. The
authoritative source is the code itself — [`piv_stats.py`](piv_stats.py) contains
every exact query. This document is the plain-language companion to it.

---

## Who built this

This tool was built by **William Thielicke, the author of PIVlab**. PIVlab ranks
first in the results. The relevant facts, so you can form your own opinion:

- The **same metric is applied to every package** (defined below); PIVlab's query
  uses the same rules as every other tool's.
- The **complete source code is public**, so every query can be inspected and
  re-run by anyone:
  <https://github.com/Shrediquette/PIV_Software_Usage_statistics>
- Every query, and the reason for it, is documented below and in the report.

Questions or suggested corrections: please open an issue.

---

## The unified metric

Every software package is counted the same way: a full-text search of the
academic literature via the [OpenAlex](https://openalex.org) API.

A paper counts for a given software if **both** of these appear in its full text
(title + abstract + body, where indexed):

1. The **software name**, as an exact quoted phrase (quotes prevent stemming —
   e.g. `"Flownizer"` does not match the word "flown").
2. The literal phrase **`"particle image velocimetry"`**.

Requiring the PIV phrase guarantees the paper is about PIV without relying on
machine-learning topic tags, and it removes false positives from names that mean
other things in other fields (e.g. *DaVis*, *TSI*, *GPIV*, *PRANA*).

Example filter (LaVision DaVis):

```
default.search:"LaVision" AND "DaVis" AND "particle image velocimetry",
publication_year:2010-2025
```

### Rules that apply equally to all packages

- **Spelling / spacing variants** are OR-combined when empirically justified, e.g.
  `("OpenPIV" OR "open piv")`. OpenAlex treats hyphens as spaces inside quotes, so
  `"PIV-lab"` already equals `"PIV lab"`.
- **Commercial tools** additionally require the **company name** in the text
  (e.g. `"Dantec" AND "Dynamic Studio"`), because commercial software is usually
  introduced together with its vendor.
- **Ambiguous short names** require a disambiguating token — an author surname for
  academic tools, a company name for commercial tools — exactly as a careful human
  reviewer would.
- Only **fully completed years (2010–last full year)** are shown.

---

## Empirical false-positive control

Variants and disambiguators are not guessed — each candidate query was tested
against OpenAlex by counting the papers it adds and inspecting a sample of titles.
A variant is only kept if it adds *genuine* papers without pulling in unrelated
ones. Some real decisions made this way:

| Case | Problem | Decision |
|---|---|---|
| **mpiv** (Mori & Chang) | Bare `"mpiv"` returned 120 papers, but most matched **"μPIV" / micro-PIV** in microfluidics and blood-flow studies — not the toolbox. | Require both author surnames: `"mpiv" AND "Mori" AND "Chang"` → 35 genuine hits. |
| **Flownizer** | Adding the camera brand "Ditect" inflated the count ~10×, because Ditect cameras are used regardless of which PIV software processes the images. | Use `"Flownizer"` only (the camera brand is excluded). |
| **PaIRS-UniNa** | Unquoted "PaIRS" matches the common English word "pairs" (>10 000 papers). | Require the exact phrase `"PaIRS-UniNa"`. |
| **PIVsuite** | The two-word "PIV suite" matches generic phrasing like "suite of PIV tools". | Use `"PIVsuite"` only. |
| **GPIV / PRANA** | "GPIV" appears in medical literature; "prana" is a Sanskrit/yoga term. | The required PIV phrase removes both. |

The same procedure was applied to PIVlab itself: `("PIVlab" OR "PIV lab")` was
adopted only after confirming "PIV lab" co-occurs with genuine PIVlab usage, and
relaxing the requirement to plain "PIV" was **rejected** because it produced too
many false positives.

---

## Known limitations

- **Commercial tools may be undercounted.** Open-source tools are almost always
  named explicitly by authors; commercial tools are sometimes described only as
  "commercial software" without a name. Full-text search cannot recover those.
- **Full-text indexing is incomplete.** OpenAlex covers ~96 % of the literature,
  but not every paper body is searchable; some mentions are missed for everyone.
- **Mentions ≠ endorsements.** A review or benchmarking paper that lists many tools
  counts for all of them equally. This affects every package the same way.
- **Journal-impact proxy.** The "Journal Quality" chart uses OpenAlex's
  `2yr_mean_citedness`, an open stand-in for the journal impact factor — not the
  official Clarivate JIF.

---

## Reproduce it yourself

The whole pipeline is one Python script with no private data:

```bash
pip install -r requirements.txt
python piv_stats.py      # writes output/piv_report.html and output/piv_data.xlsx
```

Delete the `cache/` folder to force a fresh fetch. The live report is rebuilt
automatically on the 1st of each month by GitHub Actions.

---

## Build a similar comparison for *your* field (reusable AI prompt)

This project was developed with the help of an AI coding assistant. If you want to
build an equivalent usage comparison for a different set of tools (CFD solvers,
microscopy packages, statistics software, …), you do **not** need this exact code
— you can paste the prompt below into your preferred AI tool and adapt the
bracketed parts. It encodes the same rules used here.

> I want to compare how often a set of **[software tools / methods]** are mentioned
> in the academic literature, as a reproducible usage statistic. Please write a
> single Python script that uses the **OpenAlex API**
> (free, no key; identify with a `mailto` email in the polite pool) and follows
> these rules:
>
> 1. **One identical metric for every tool.** For each tool, count works where the
>    full text (via `default.search`) contains the tool's name **as a quoted exact
>    phrase** AND a required disambiguating context phrase
>    **["<the field's defining phrase, e.g. 'particle image velocimetry'>"]**.
>    Never give one tool a looser rule than another.
> 2. **Control false positives empirically.** For any short, generic, or ambiguous
>    name, add a disambiguating token (an author surname for academic tools, the
>    company name for commercial ones). Before committing a query, fetch the count
>    and a sample of titles and check they are genuinely on-topic. Prefer missing a
>    few papers over including false positives — and apply this tightening to *all*
>    tools, including any you might personally favour.
> 3. **Handle spelling/spacing variants** with OR-combined quoted phrases, but only
>    keep a variant if it adds real papers without adding noise.
> 4. **Restrict to fully completed years** and cache every API response to one JSON
>    file per query so re-runs are fast and reproducible.
> 5. Produce: a per-year trend, total counts, market share, an open-source vs.
>    commercial split, field and country distributions, and an interactive HTML
>    report. Also export the raw numbers to a spreadsheet.
> 6. **Disclose any conflict of interest** in the output, and document every query
>    so the method can be audited.
>
> Here is my list of tools with their canonical names, vendors/authors, and known
> spelling variants: **[list them]**.

Adapting this for a new domain mainly means changing the required context phrase
and the tool list — the rules stay the same.
