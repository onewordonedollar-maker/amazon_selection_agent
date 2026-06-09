# Agent Instructions

Read these files before making changes:

1. `README.md`
2. `docs/lessons-learned/project-guardrails.md`
3. `docs/lessons-learned/pitfalls.md`
4. The tests related to the requested behavior

## Required Workflow

- Inspect the existing implementation before editing.
- Handle one clearly scoped problem at a time.
- Verify each fix before moving to the next problem.
- Run `python -m unittest discover -s tests -v` before delivery.
- Use the real local browser for frontend and collection-flow verification.
- Preserve unrelated user changes.

## Protected Behavior

Do not change Amazon page waiting, scrolling, pagination, SellerSprite loading,
ASIN deduplication, or filter semantics without explicit user approval.

Before proposing such a change, write:

```text
★ 涉及核心采集/筛选规则变更
```

Then explain the reason, impact, tests, and rollback plan. Wait for approval.

## Data Safety

Never commit:

- `chrome_profile/`
- Credentials or browser login state
- Runtime collection results
- Temporary checkpoints and caches

The validated category mapping at
`outputs/category_links_learned.json` is intentionally versioned.

