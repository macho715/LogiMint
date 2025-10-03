# Plan

## Goal
Refactor and upgrade the HVDC logistics scripts into a modular package with shared utilities, centralized configuration, robust error handling, and performance improvements while keeping the CLI workflow stable.

## TDD Cycle Overview
1. **RED**: Add or update tests capturing expected behaviors for scanning, mapping, and Excel generation with new utilities.
2. **GREEN**: Implement utilities, configs, and refactors to satisfy tests while preserving external contracts.
3. **REFACTOR**: Optimize structure, improve error handling, add async/pipeline enhancements, and ensure code quality gates pass.

## Iterative Steps

### Iteration 1 – Scanning utilities foundation
- Add fixtures and tests for email scanning edge cases (empty folder, Outlook exclusion, non-ASCII). *(RED)*
- Create `utils/file_handler.py` and `utils/email_parser.py` with the minimal APIs to satisfy the tests. *(GREEN)*
- Refine implementations for streaming IO and generators; add custom exceptions. *(REFACTOR)*

### Iteration 2 – Pattern matching consolidation
- Introduce tests for regex normalization/fuzzy matching with data-driven cases. *(RED)*
- Implement `utils/pattern_matcher.py` and centralize regex constants under `config/patterns.py`. *(GREEN)*
- Remove duplication in existing scripts by adopting the shared module. *(REFACTOR)*

### Iteration 3 – Config and logging standardization
- Write tests validating configuration defaults and logger factory behavior. *(RED)*
- Add `config/settings.py`, structured logging helpers, and integrate into refactored modules. *(GREEN)*
- Ensure scripts use injected settings/CLI overrides and rotating handlers. *(REFACTOR)*

### Iteration 4 – Priority script refactors
- Capture regression tests for `comprehensive_email_mapper`, `folder_title_mapper`, `hvdc_cargo_tracking_system` covering main flows. *(RED)*
- Split large modules into orchestrators leveraging new utilities and async batches. *(GREEN)*
- Polish error handling, docstrings, and type hints; ensure LOC targets met. *(REFACTOR)*

### Iteration 5 – Pipeline integration and Excel stability
- Add integration test for scan→map→excel pipeline ensuring CLI compatibility. *(RED)*
- Update scripts to use streaming writes, chunked batching, and new config toggles. *(GREEN)*
- Tune performance, update documentation, changelog, and finalize logging/metrics. *(REFACTOR)*

### Iteration 6 – Quality gates & docs
- Run `pytest`, coverage, `black --check`, `isort --check-only`, `flake8`, `mypy --strict`. *(GREEN)*
- Document verification metrics, update README runbook and CHANGELOG, highlight residual risks. *(REFACTOR)*

