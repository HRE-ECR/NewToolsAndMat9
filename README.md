# PdfTableExtractor (v10)

Repo-ready WPF (.NET 8) app that builds a self-contained single-file EXE using GitHub Actions.

## MaterialDesignThemes note
MaterialDesignThemes v5+ renamed `MaterialDesignTheme.Defaults.xaml` to `MaterialDesign2.Defaults.xaml` / `MaterialDesign3.Defaults.xaml`.
This repo uses MaterialDesign2 + ObsoleteBrushes mapping so older brush keys like `MaterialDesignPaper` still work.

## Workflow
- Triggers on push to `main` and also supports manual runs via `workflow_dispatch`.
- Uploads the entire publish folder as an artifact.
