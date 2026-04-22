# tree-sitter-vba

Tree-sitter grammar for Visual Basic for Applications (VBA).

Consumed by the Zed extension [`tmepple/zed-vba`](https://github.com/tmepple/zed-vba), which pins this grammar by commit SHA in its `extension.toml`.

## Setup

```bash
npm install
```

Installs `tree-sitter-cli` and the node bindings.

## Workflow

Edit `grammar.js`, then regenerate and test:

```bash
npx tree-sitter generate     # regenerate src/parser.c from grammar.js
npx tree-sitter test         # run corpus tests in test/corpus/
npx tree-sitter parse <file> # parse a file and print the tree
```

Add new test cases to `test/corpus/` — one `.txt` file per feature area, following the tree-sitter corpus format.

## Releasing a new grammar version

The Zed extension pins this grammar by commit SHA, not by tag, so shipping a grammar change is just:

1. Commit and push changes to `master`.
2. Grab the new SHA: `git rev-parse origin/master`.
3. In the `zed-vba` repo, update `rev = "<sha>"` under `[grammars.vba]` in `extension.toml`.
4. Commit + push `zed-vba`.
5. Users reinstall the dev extension (or get the update via the Zed registry once published).

If you bump the `version` in `tree-sitter.json` / `package.json`, do it in the same commit as the grammar change so the SHA captures both.

## Layout

- `grammar.js` — grammar definition (source of truth)
- `src/parser.c`, `src/tree_sitter/` — generated parser; regenerate with `tree-sitter generate`
- `test/corpus/` — test cases
- `queries/` — example highlight queries (the Zed-specific ones live in `zed-vba/languages/vba/`)
- `bindings/` — language bindings (node, rust, etc.) for non-Zed consumers
