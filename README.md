<img src="docs/eyecatch-banner.png" alt="xlreset" width="100%">

<h3 align="center"><em>xlreset is a zero-dependency CLI that resets every sheet back to A1, 100% zoom.</em></h3>

| Before | After |
| --- | --- |
| ![Before: scrolled, zoomed, wrong cell selected](docs/screenshot-before.png) | ![After: A1 selected at 100% zoom](docs/screenshot-after.png) |

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
- the active tab to the first sheet

Everything else is left untouched.

xlreset is a CLI-only tool with no configuration flags. Feature requests?
Please [open a GitHub issue](https://github.com/RKasai127/xlreset/issues).

## License

MIT
