# Ship Duplicate Review MVP
## Progress Report

**Prepared by:** Mohamed Elsabah  
**Project context:** Historical ship-ledger duplicate review  
**Date:** May 10, 2026

---

## Executive Summary
Historical ship-ledger datasets frequently contain inconsistent vessel-name spellings, punctuation variations, capitalization differences, and transcription uncertainty. These inconsistencies make duplicate detection difficult and can introduce research errors when records are merged without sufficient evidence.

The goal of this MVP is to support historians and reviewers in identifying possible duplicate vessel names while preserving research integrity. The application is designed as a human-in-the-loop review system, where the software proposes possible matches and the reviewer makes final decisions.

This design emphasizes three principles: auditable decisions, reversible actions, and evidence-first review. Rather than blindly standardizing names, the system separates safer cases from uncertain cases and keeps a traceable decision history.

## Problem Background
Blind automatic merging is risky in historical datasets because name similarity does not always imply identity equivalence. A vessel may be referenced with multiple spellings, but similar names can also refer to different vessels in different periods or cargo contexts.

Historians need evidence before deciding that two records should be merged. For this reason, the MVP surfaces supporting signals (name similarity and context evidence) and lets reviewers decide among merge, keep separate, or unsure.

The workflow intentionally separates obvious formatting duplicates from uncertain historical identity cases. This prioritization helps reviewers focus their effort on difficult decisions while still benefiting from automation in low-risk scenarios.

## Current MVP Workflow
The current end-to-end workflow is:

1. Upload workbook (`.xlsx` or `.xls`).
2. Choose worksheet.
3. Map primary/entity column and optional evidence columns.
4. Generate safe auto-merge groups.
5. Generate fuzzy manual-review candidates.
6. Review candidates as `merge`, `keep separate`, or `unsure`.
7. Inspect merge history.
8. Undo saved decisions.
9. Save/load review session JSON.
10. Export reviewed workbook.
11. Export standardized workbook.
12. Export audit CSVs.

## Current Algorithm Baseline
The MVP algorithm currently follows a pragmatic evidence-scoring approach:

- **Name normalization:** Names are cleaned to reduce punctuation/case/spaces noise.
- **Strict-key safe auto-groups:** Obvious formatting variants are grouped using strict normalized keys.
- **Fuzzy matching:** RapidFuzz-based similarity is used to generate ambiguous candidate pairs.
- **Weighted evidence scoring:** Candidate ranking uses weighted signals (name similarity plus year, unit, type/category, and amount evidence).
- **Human final decision:** The reviewer approves or rejects merges; the algorithm does not finalize uncertain merges on its own.

At this stage, the MVP does **not** use K-means clustering or trained machine-learning models.

## Algorithm Direction After Meeting
Current direction is to defer K-means/ML-style clustering in the short term. The reasons are:

- The app must remain responsive for practical review sessions.
- Reliable ML direction requires real reviewer criteria and labeled decisions first.

Edit-distance concepts from analysis/design material remain useful for understanding insertions, deletions, and substitutions in messy names. However, for the current MVP, optimized fuzzy matching is more practical and implementation-ready for noisy historical strings.

The immediate research direction is to observe/interview reviewers and capture how they actually decide merges, then use those criteria to guide later model evolution.

## Performance and Speed Baseline (Discussion)
Performance is an important concern because large workbooks can slow interactive review. This report documents what should be measured next, without introducing fabricated timing values.

Planned measurement points:

- Workbook load time
- Preprocessing time
- Name-statistics generation time
- Safe auto-group generation time
- Manual candidate generation time
- Merge-output generation time
- Export generation time

Expected bottleneck hypothesis: manual candidate generation is likely to be the dominant cost in larger datasets.

### Future Timing Table Template
| Stage | Metric Definition | Dataset Identifier | Run Date | Measured Time (s) | Notes |
|---|---|---|---|---:|---|
| Workbook load | Read workbook + selected sheet | _TBD_ | _TBD_ | _TBD_ | |
| Preprocessing | Row cleaning + normalized fields | _TBD_ | _TBD_ | _TBD_ | |
| Name statistics | Aggregation of name-level stats | _TBD_ | _TBD_ | _TBD_ | |
| Safe auto-groups | Strict-key grouping stage | _TBD_ | _TBD_ | _TBD_ | |
| Candidate generation | Fuzzy pair generation + scoring | _TBD_ | _TBD_ | _TBD_ | |
| Merge outputs | History + canonical mapping build | _TBD_ | _TBD_ | _TBD_ | |
| Export generation | Reviewed/standardized workbook creation | _TBD_ | _TBD_ | _TBD_ | |

## UX/User Stories
- As a historian, I want to see why two vessel names were suggested as a possible match so that I can decide whether to merge them confidently.
- As a reviewer, I want obvious formatting duplicates separated from uncertain historical identity cases so that I can focus on difficult decisions.
- As a researcher, I want saved review decisions so that future algorithm changes can be evaluated against real human-reviewed examples.
- As a user, I want undo and session save/load so that I can safely review over multiple sittings.

## Progress Over Time
Key MVP progress milestones include:

- Initial duplicate-review MVP
- Generic Excel column mapping
- Safe auto-merges
- Manual review queue
- Merge history and undo
- Cleaned workbook export
- Standardized workbook export
- Excel formatting preservation
- Session JSON save/load
- Hardened session validation
- UI improvements

## Screenshot Placeholders
> Placeholder only. Screenshots will be inserted manually.

- **Screenshot 1:** Workbook upload and column mapping
- **Screenshot 2:** Safe auto-merge review
- **Screenshot 3:** Manual review queue
- **Screenshot 4:** Merge history and undo
- **Screenshot 5:** Export and session controls
- **Screenshot 6:** Reviewed workbook output

## Next Steps
1. Observe/interview reviewers during real dedupe sessions.
2. Document reviewer merge criteria and evidence preferences.
3. Create a formal speed baseline from measured runs.
4. Improve speed after baseline measurements are available.
5. Refine UX based on historian feedback.
6. Consider ML approaches only after enough labeled review decisions exist.
7. Prepare demo screenshots and a sample dataset for presentations.
