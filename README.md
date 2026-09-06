# xlreset
CLI to set Excel's initial view to cell A1 and 100% zoom.

## Install

```sh
pipx install xlreset
```

## Usage

```sh
xlreset file.xlsx
```

The file is rewritten **in place** (no backup is created). For every sheet in
the given `.xlsx`/`.xlsm` file, xlreset sets:

- scroll position (`topLeftCell`) to `A1`
- zoom to 100% (`zoomScale` and `zoomScaleNormal`)
- the active cell/selection (`activeCell`, `sqref`) to `A1`

It leaves untouched: frozen panes, sheet protection, macros (`vbaProject.bin`
in `.xlsm` files), and everything else in the workbook.

xlreset is a CLI-only tool with no configuration flags in this version.
Feature requests should be filed as
[GitHub issues](https://github.com/RKasai127/xlreset/issues) rather than
expected as flags.

## License

MIT
