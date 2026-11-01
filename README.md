# PIV Software Usage Statistics

An automated analysis of how Particle Image Velocimetry (PIV) software packages 
are cited in academic publications, based on data from [OpenAlex](https://openalex.org).
The methodology is shown here:
[Methodology](https://github.com/Shrediquette/PIV_Software_Usage_statistics/blob/main/METHODOLOGY.md)

## What this shows

- Number of academic papers mentioning each PIV software package (2010–present)
- Trends over time, market share, and open source vs. commercial breakdown
- Research fields and countries where each software is used
- Journal quality indicator (weighted mean impact per software)

## Software covered

32 packages across three categories: Open Source, Free / Academic, and Commercial —
including PIVlab, OpenPIV, MatPIV, LaVision DaVis, Dantec Dynamic Studio, and more.

## Methodology

Papers are identified via the OpenAlex full-text search API using exact phrase 
matching (software name + "particle image velocimetry"). Counts reflect papers that 
explicitly mention the software, not just any PIV paper.

## Live report

https://shrediquette.github.io/PIV_Software_Usage_statistics/

Updated automatically on the 1st of every month via GitHub Actions.

## Local usage

```bash
pip install -r requirements.txt
export OPENALEX_API_KEY=your_key_here   # free key: https://openalex.org/settings/api
python piv_stats.py
```

The script needs a free OpenAlex API key in the `OPENALEX_API_KEY` environment
variable (on Windows: `set OPENALEX_API_KEY=your_key_here`). For the monthly
GitHub Actions run, store it as a repository secret with the same name
(Settings → Secrets and variables → Actions). If OpenAlex refuses requests or
data comes back incomplete, the run stops without overwriting the published report.
