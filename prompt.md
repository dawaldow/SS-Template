# Purpose
Read the latest Smartsheet project data and produce a factual, action-oriented daily summary that I can as the PM  can scan in 30–60 seconds.

# Sources and Reporting Date
Use the current contents of:
1. `project_summary.docx` for project name, Generated timestamp, executive context, portfolio counts, and summary exceptions.
2. `project_data.xlsx` for WBS number, owners, status, dates, completion, forecast, variance, risks, dependencies, workstreams, milestones, and Smartsheet links.

If another workbook is supplied, analyze its current data and any explicitly dated prior snapshot. Never carry values forward unless supported by current files.

Use the **Generated date and time in `project_summary.docx`** as the report date. Never substitute the system date, modified date, project start date, or dates from messages. Show the supplied timezone; otherwise write **timezone not specified**.

Use organizational information only to resolve owners or documented decisions/dependencies. Mark employment status **Unverified** unless reliable organizational data verifies it; never infer inactivity from missing profile data.

Always provide the task name and wbs number together for each reference or occurrence of a task.

# Latest File Selection (Critical)

Always identify all matching files named:
- project_summary.docx
- project_data.xlsx

If multiple versions exist, select ONLY the files with the most recent Modified timestamp.

Never use an older version when a newer version exists.

Before analysis, explicitly validate:
1. project_summary.docx Modified timestamp
2. project_data.xlsx Modified timestamp
3. Generated timestamp inside project_summary.docx

If a newer matching file is discovered after an initial retrieval, discard all prior calculations and restart analysis using the newest files.

The report MUST be based on the latest available export, not the first search result returned.

Report the selected file timestamps in the Snapshot column.

# Analysis Steps
1. Confirm both primary files are accessible, refer to the same project, and contain parseable data.
2. Compare overlapping counts and explicit timestamps. Report stale, partial, inconsistent, missing, or unparseable data; never silently reconcile conflicts.
3. If a comparable prior snapshot exists, report only material changes in status, completion, dates, milestones, ownership, blockers, dependencies, risk, or forecast confidence. Note incompatible populations or definitions. If unavailable, use one concise **Comparison unavailable** row.
4. Identify material exceptions: blocked work, overdue items with downstream impact, at-risk or behind-schedule work, incomplete tasks due within 7 days, milestones due within 30 days, ownership gaps, schedule variance, concentrated dependencies, and data inconsistencies.
5. Report completed tasks only when completion closed a risk, removed a blocker, released a dependency, changed a milestone forecast, resolved overdue work, or materially affects a decision.
6. For each schedule concern, give the task/milestone, supporting fact, owner, due date, and documented downstream effect in one sentence.
7. Rank actions by urgency, downstream impact, due date, work unblocked, and need for leadership intervention.

Only call work **blocked**, **critical**, or on the **critical path** when explicitly stated or directly proven by dependency data. Label supported analysis not stored as a source fact with **Potential risk:**. Never invent dates, owners, blockers, dependencies, changes, explanations, or comparison values.

# Overall Status
- **🔴 Act Now:** A critical milestone is overdue, a major dependency is blocked, a committed date requires leadership intervention, or an immediate executive decision is explicitly required. A lone noncritical overdue task is insufficient.
- **🟡 Watch:** A material risk or exception needs active management, but recovery appears feasible.
- **🟢 On Track:** No material exception threatens committed dates, no major dependency is blocked, and risks are controlled without executive intervention.

Use the most severe supported condition across schedule, scope, budget, resources, and dependencies; do not average indicators.

# Output Format
Use Markdown headings and compact tables suitable for Teams. Keep cells short, paragraphs to two sentences, and the report near one Teams screen when findings permit. Omit empty sections and unavailable metrics. Preserve source names, WBS, owners, dates, statuses, and links. Avoid raw HTML, copied grids, filler, duplicated facts, a narrative executive summary, and generic closing comments.

## Project Daily Summary
| Project | Report date | Snapshot | Overall status |
|---|---|---|---|
| [Linked project] | [Generated timestamp] ([timezone]) | [Current/stale/partial/inconsistent + source note] | [🟢 On Track/🟡 Watch/🔴 Act Now] |

## Today’s Action Required
| Most Important Decision or Action Today |
|---|
| **[Complete verb-led sentence stating the required action or decision, one owner, required timing, and the outcome protected, resolved, or unblocked.]** |

Write exactly one complete sentence in this section. It must contain a specific action or decision, owner, timing, and intended outcome; never use a fragment, placeholder, or unfinished thought. If it references one or more tasks, add rows listing each task name, WBS number, status, start date, end date, and predecessors so the action has complete supporting context.

## Dashboard
Include only decision-useful available KPIs: overall completion, overdue, blocked, at-risk, due within 7 days, due within 30 days, behind schedule, unassigned leaf tasks, and at-risk workstreams.

| KPI | Current | Change | Executive interpretation |
|---|---:|---:|---|
| [KPI] | [Value] | [Change/No change/No baseline] | [Concise impact] |

Use **No baseline** when comparison is unavailable. Count incomplete due-soon work separately from completed work.

## Significant Changes
Include only material changes or a comparison limitation; omit when a valid comparison shows none.

| Area | Previous | Current | Material impact |
|---|---|---|---|
| [Task/milestone/risk/owner/metric] | [Prior or Comparison unavailable] | [Current] | [Why it matters] |

## Priority Actions
Include no more than five exceptions, ranked by executive importance.  Ensure that the  exceptions are in a verb-led action format. Each must name one owner, due date or urgency, and what it unblocks. Use **Owner needed** when none is supported.

| # | Action | Owner | Due / urgency | Impact |
|---:|---|---|---|---|
| 1 | [Verb-led action] | [Owner/Owner needed] | [Date/Today/Overdue] | [Result] |

## Upcoming Milestones
Include only milestones due within 30 days.

| Milestone | Owner | Target date | Complete | Status / confidence |
|---|---|---:|---:|---|
| [Milestone] | [Owner] | [Date] | [%] | [🟢 On Track/🟡 Watch/🔴 At Risk] |

# Final Validation
Before responding, verify that **Today’s Action Required** is one grammatically complete sentence containing an action or decision, owner, timing, and intended outcome. Confirm there are no fragments, placeholders, unfinished Markdown, or unmatched table delimiters. If space is constrained, shorten or omit lower-priority sections before shortening this required action.

# Error Handling
- Missing Generated timestamp: state the authoritative date is unavailable, continue supported analysis, and report the gap; do not substitute a date.
- One source unavailable: identify it, limit findings to available evidence, and report the limitation.
- Missing columns or unparseable values: identify fields and affected IDs/rows, exclude unsupported calculations, and continue supported analysis.
- Conflicting sources: show each value and its source; apply the stated source responsibilities only where they resolve the field.
- Neither primary source available: identify both files and state that the daily report cannot be produced; do not create a speculative status.