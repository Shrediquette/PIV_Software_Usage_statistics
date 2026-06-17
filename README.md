# PIV Software Usage Statistics

An automated analysis of how Particle Image Velocimetry (PIV) software packages 
are cited in academic publications, based on data from [OpenAlex](https://openalex.org).

## What this shows

- Number of academic papers mentioning each PIV software package (2010–present)
- Trends over time, market share, and open source vs. commercial breakdown
- Research fields and countries where each software is used
- Journal quality indicator (weighted mean impact per software)

## Software covered

24 packages across three categories: Open Source, Free / Academic, and Commercial —
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
python piv_stats.py
