# Story Info + Vocab & Click Words

Separate Streamlit app that extends the existing Vocab & Click Words workflow.

## Workflow

This app is split into two stages.

### Stage 1. Story Info

Download `Story_Confirmed_Template.xlsx` in the app and fill:

- `ID`
- `Title`
- `Level`
- `Base Text`

`Level` accepts `1-4` or `Pre A1`, `A1`, `A2`, `B1`, `B2`, `C1`, `C2`.

Stage 1 creates:

- estimated CEFR and Lexile from `Base Text`
- word count and scene count
- category, book mood, summary, keywords, intro script
- `Easy Version`
- `Difficult Version`

The download file is `Story_Info_Result.xlsx`. It contains:

- `Story_Info`
- `Vocab_Input`

### Stage 2. Vocab & Click Words

Stage 2 requires all three text versions:

- `Normal Ver.`
- `Easy Ver.`
- `Difficult Ver.`

You can use the current Stage 1 session result directly, or upload `Story_Info_Result.xlsx` later. The app reads the `Vocab_Input` sheet when it exists.

The download file is `Vocab_Click_Words_Analysis.xlsx`.

## Level Logic

- Input `Level` is used only for generating `Easy Version` and `Difficult Version`.
- `Detected CEFR` and `Lexile` are estimated from `Base Text`.
- Vocab & Click Words filtering uses the estimated `Detected CEFR` with `lcms_cefr.csv`.

Both stages support checkpoint JSON files for longer batches.
