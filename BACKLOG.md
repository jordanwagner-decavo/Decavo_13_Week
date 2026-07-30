# Backlog / Open Design Questions

## 2026-07-30 — SO Total "past due" (to push) definition

**Status:** Deferred — needs a decision before implementing.

### Context
The **SO TOTALS** column now shows total open qty on order (from P&P) plus up to
two warning lines:
- `⚠ N to plan` — open on-order qty not covered by the plan.
- `↩ N to push` — currently past-due *plan-vs-actual* miss.

### The problem we found
The `↩ N to push` line is (almost) always empty even though there are many
genuinely past-due items. Root cause: today it is computed from **`rowMiss`** =
for each past week, `planned − actual` (qty we scheduled in the 13-week but did
not ship). It never references the P&P/SO due dates. It reads 0 whenever:
- the qty was never put in a *past* 13-week column (nothing planned → nothing to
  "miss"), or
- we shipped whatever we had planned.

So it measures **execution vs. our own plan**, not **lateness vs. the customer's
SO due date** — which is what we actually want to surface.

### Two distinct metrics (keep them straight)
1. **Late vs. SO** (what we want for "past due"): open P&P qty whose **due date
   is already in the past**. Real customer lateness. Data is available — each
   demand line carries `{ due, qty }`.
2. **Plan-vs-actual miss** (`rowMiss`, current behavior): did we execute our own
   plan? Still meaningful, but it is NOT "past due."

### Reconcile SO (the per-week view)
`Reconcile SO` already does the per-week version of metric #1: it buckets P&P
open qty by due date and flags weeks where cumulative demand outruns cumulative
supply (`behindWeeks`). That is the "planned dates vs SO due dates" visibility at
the cell level. **SO Total = the rollup number; Reconcile SO = the week-by-week
picture** — they complement each other.

- **Known inconsistency to fix:** `behindWeeks` currently counts **past shipped
  actuals** as supply. Since P&P Open Qty already excludes shipped units, this
  makes Reconcile *under*-flag lateness. Fix: count only forward planned qty
  (exclude past actuals) so the red cells match reality.

### OPEN QUESTION — how to define the SO Total "past due" line
Pick one (using P&P due dates):
- **(a) Late, raw** — open P&P qty whose due date is already past; counts ALL late
  qty regardless of any catch-up plan; clears only when shipped. Simplest.
- **(b) Late, unrecovered** — past-due open qty MINUS what we've since planned in
  current/future weeks to catch up; drops to 0 once recovery is scheduled. More
  actionable, more complex (cumulative allocation like `behindWeeks`).
- **(c) Planned but late** — only qty that IS planned but scheduled in/after a week
  already past its SO due date. Keeps `to plan` (unplanned) and `to push`
  (planned-but-late) strictly disjoint.

### Secondary questions to resolve alongside
- Should `to plan` and `to push` be **disjoint** or allowed to **overlap**?
- Should `coverage` keep including the past-due execution miss (`rowMiss`), or
  switch to forward-planned-only now that Open Qty excludes shipped?
- Confirm the "past due" cutoff: due date before **today**, or before the
  **current week's start/end** (align with `behindWeeks` week bucketing).
- Label/wording: keep "to push" vs. rename to "past due"?

### Relevant code
- `13-WK Front End.html`: `renderTotal`, `coverage`, `rowMiss`, `onOrder`,
  `behindWeeks`, `applyReconcile`.
- `13-WK Back End.gs`: `readPP_` (builds `demand` with `due`/`qty`), `PP` config
  (SO=B, DPN=F, Open Qty=H, Due=J).
