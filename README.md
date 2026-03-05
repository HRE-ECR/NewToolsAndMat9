# PdfTableExtractor (v9)

## Fix for crash: missing MaterialDesignTheme.Defaults.xaml
MaterialDesignThemes v5+ renamed the defaults dictionary.
This repo now references `MaterialDesign2.Defaults.xaml` and includes `MaterialDesignTheme.ObsoleteBrushes.xaml` for compatibility.

If the app still closes, check `crash.log` and `app.log` next to the EXE.
