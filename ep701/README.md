# EP701 bio-fact merge

Drop these files into this folder, then run the script:

```
EP701_personalised.csv   master (701 firms)
bios_v3.csv              fresh bio text (287 firms)
```

```bash
pip install pandas openpyxl
python3 ep701/ep701_bio_merge.py            # reads/writes in ep701/
python3 ep701/ep701_bio_merge.py --dir /path/to/folder
```

Outputs written next to the inputs:

- `EP701_personalised_v3.csv`
- `EP701_personalised_v3.xlsx` (All sheet plus one sheet per tier, bold header, frozen header row)
- `EP701_v3_report.md` (B to D, C to D, still C and why, personal swaps, duplicates, icp flags)

What it does, in order:

1. Keeps only bios_v3 rows with a blank `error` and `bio_text` over 300 characters.
2. Pulls one professional fact from `bio_text` only (licensed since year, founded year, board
   certified, wrote a book, prior career, bar role, radio show). Any candidate whose quote or
   line touches family, kids, hometown, pets or health is skipped and the next fact is tried.
3. `bio_quote` is always a slice of the bio text (max 20 words) and is re-checked verbatim.
   `bio_fact` is max 12 words. `first_line` is max 18 words, no dashes, and starts with
   Saw / Noticed / Read / Came across. Openers rotate so lines do not all look alike.
4. Merges on normalised domain. `B_events_page_only` rows keep the old line in `seminar_line`
   and become `D_bio_fact`; `C_none` rows become `D_bio_fact`. `A_dated_event` and existing
   `D_bio_fact` rows are never modified. Rows with no usable fact keep their tier and get the
   reason in `bio_status` (`NO_BIO_v3`, `BIO_ERROR_v3`, `BIO_SHORT_v3`, `NO_FACT_FOUND_v3`,
   `FACT_PERSONAL_ONLY_v3`, `QUOTE_NOT_VERBATIM_v3`).
5. QA over every row into `qa_flags`: word count 5 to 20, dashes, opener, personal words,
   duplicate first_lines across domains, quote still present in bio. `icp_flag` marks firms whose
   text reads as injury, mass_tort or criminal rather than estate planning.

Nothing is invented. Rows the regex extractor cannot handle are listed under "Still C_none" in
the report for a manual pass.
