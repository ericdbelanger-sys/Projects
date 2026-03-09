# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running the Calculator

```bash
python calculator.py
```

## Architecture

`calculator.py` is a single-file CLI calculator with three layers:

1. **Input preprocessing** (`insert_implicit_mul`) — regex-based pass that rewrites implicit multiplication (e.g. `3(4+2)` → `3*(4+2)`) before parsing.
2. **Safe expression evaluation** (`eval_expr`) — walks a Python `ast` tree, supporting `+`, `-`, `*`, `/`, `**`, unary negation, and parentheses. Uses `ast.parse` instead of `eval` to prevent arbitrary code execution.
3. **REPL loop** (`main`) — reads expressions, calls `calculate`, and formats results (whole-number floats are displayed as integers).
