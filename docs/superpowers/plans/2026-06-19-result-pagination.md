# Result Pagination Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add pagination to post-collection result browsing without changing collection, SellerSprite loading, ASIN deduplication, or filter semantics.

**Architecture:** Put pagination math in a small pure helper module and keep Streamlit responsible only for state, controls, and slicing the already-filtered product list. Selection is still stored on each `Product` and widget keys by ASIN, so checked products remain selected across pages.

**Tech Stack:** Python, Streamlit, existing `unittest` suite.

---

### Task 1: Pure Pagination Helpers

**Files:**
- Create: `src/result_pagination.py`
- Test: `tests/test_result_pagination.py`

- [ ] **Step 1: Write failing tests** for default page size, page bounds, range labels, and page slicing.
- [ ] **Step 2: Run targeted tests** with `.runtime\python\python.exe -m unittest tests.test_result_pagination -v` and confirm they fail because the helper module does not exist.
- [ ] **Step 3: Implement helpers**: `PAGE_SIZE_OPTIONS`, `normalize_page_size`, `page_count`, `clamp_page`, `page_slice`, and `page_range_label`.
- [ ] **Step 4: Run targeted tests** and confirm they pass.

### Task 2: Streamlit Result Pagination UI

**Files:**
- Modify: `streamlit_app.py`
- Test: `tests/test_streamlit_ui_state.py`

- [ ] **Step 1: Add source-level tests** confirming pagination state keys, toolbar controls, product slicing, and export-range labels exist.
- [ ] **Step 2: Run targeted tests** and confirm they fail before wiring UI.
- [ ] **Step 3: Add Streamlit state and callbacks** for page size, current page, export scope, and page reset after filter/load/collection changes.
- [ ] **Step 4: Render top and bottom pagination controls** with page size `50 / 100 / 200`, previous/next buttons, page number input, and range text.
- [ ] **Step 5: Slice product cards and table rows** using current page products only while keeping summary metrics and selected counts based on all filtered products.
- [ ] **Step 6: Keep selection cross-page** by continuing to use ASIN-based checkbox keys.
- [ ] **Step 7: Add export scope** for selected products, current page, or all current filtered results.
- [ ] **Step 8: Run targeted tests** and confirm they pass.

### Task 3: Verification

**Files:**
- No new production files.

- [ ] **Step 1: Run full tests** with `.runtime\python\python.exe -m unittest discover -s tests -v`.
- [ ] **Step 2: Verify local UI** at `http://localhost:8501/`: pagination controls appear, default page size is 50, switching to 100 works, top and bottom controls stay in sync, and no Streamlit errors appear.
- [ ] **Step 3: Check git status** and summarize changed files.
