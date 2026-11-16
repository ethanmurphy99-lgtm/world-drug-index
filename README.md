# World Drug Index

Analysis code for the **World Drug Index (WDI)** — a global composite index
measuring and comparing substance-use outcomes across countries, computed
overall and separately for five substance categories (cannabis,
hallucinogens, amphetamines, cocaine, opioids), each split by all/male/female
population.

This project follows the same methodological approach as the
[European Drug Index](https://doi.org/10.1016/j.drugpo.2026.105451)
(Murphy, Comiskey, Walsh & Murphy, 2026, *International Journal of Drug
Policy*), extended to a global scope. A companion publication for the World
Drug Index is in preparation; this repository will be updated with a
citation and a DOI-archived release once that work is published.

## What's here

| File | Purpose |
|---|---|
| `WDI_validation.R` | Main scoring script. Reads a country-level indicator workbook and the indicator weighting scheme, computes normalized 0–1 scores (winsorized 5th/95th-percentile benchmarking per indicator), aggregates them into composite and per-substance indices by sex, assigns data-availability tiers, and writes a results workbook. |
| `weights.xlsx` | The indicator weighting scheme: one row per indicator code, giving its composite/male/female weight, group ID (used for proportional weight redistribution when a country is missing an indicator), and substance-category code. |
| `data/` | Placeholder — see [`data/README.md`](data/README.md). The raw country-level indicator data is withheld until the World Drug Index paper is published. |

## How the scoring works (summary)

1. **Benchmarking** — for every indicator code, pool the raw values reported by every country and compute the 5th/95th percentile as the normalization floor/ceiling (winsorized, to reduce sensitivity to outliers).
2. **Normalization** — each country's raw value is rescaled to 0–1 against that floor/ceiling; indicators where a *lower* raw value is the more desirable outcome (e.g. prevalence rates) are flipped so that 1 always means "better."
3. **Weighted aggregation** — normalized indicator scores are combined into a composite score and five substance-specific sub-scores (plus all/male/female breakdowns), using the weights in `weights.xlsx`. If a country is missing an indicator within a weighting group, that indicator's weight is redistributed proportionally among the indicators the country does have data for, so each group still sums to its full intended weight.
4. **Data-availability tiers** — countries are additionally tiered (Tier-1/2/3) by how many of the relevant indicators they actually report, using the 33rd/66th percentile of indicator-count across countries as the cut points. This is a transparency flag on data coverage, not a score adjustment.

Full detail is in the comments throughout `WDI_validation.R`.

## Running it

The script currently points at local file paths (edit `INPUT_DIR` near the top of the script to point at your own copies):

- `database_input.xlsx` — one sheet per country with raw indicator values, plus a "Global Min/Max" benchmark sheet (not yet published — see `data/README.md`).
- `weights.xlsx` — included in this repo.

Requires R with the `readxl`, `openxlsx`, `dplyr`, and `stringr` packages (the script will install any that are missing).

```r
Rscript WDI_validation.R
```

Output: `WDI_Global_Results.xlsx`, written next to the input files, containing one tab per category (Total + 5 substances) with country scores and tiers, a detailed per-indicator benchmark diagnostics tab, and one detail tab per country.

## Citation

A citation (with DOI, once assigned) will be added here when the World Drug
Index publication and dataset are released. In the meantime, please cite the
methodologically related European Drug Index paper:

> Murphy, E., Comiskey, C., Walsh, P.P., Murphy, E. (2026). The European
> drug index: Measuring and comparing national substance use across the
> European Union to aid policy and planning. *International Journal of Drug
> Policy*, 156, 105451. https://doi.org/10.1016/j.drugpo.2026.105451

## License

[MIT](LICENSE).
