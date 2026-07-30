# Story Info + Vocab & Click Words

Separate Streamlit app that extends the existing Vocab & Click Words workflow.

## Workflow

This app is split into two stages.

### Stage 1. Story Info

Download `Story_Confirmed_Template.xlsx` in the app and fill:

- `ID`
- `Title`
- `Platform Level`
- `Base Text`

`Platform Level` accepts `1-4`. Legacy CEFR values such as `Pre A1`, `A1`, `A2`, `B1`, `B2`, `C1`, and `C2` are still accepted for older files.

Stage 1 creates:

- estimated CEFR and Lexile from `Base Text`
- `Flagged Words` that exceed the CEFR ceiling mapped from `Platform Level`, with replacement suggestions when available
- word count and scene count
- three ordered categories, three ordered book moods, summary, keywords, intro script, movie book script
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

You can also start from Stage 2 only with `Vocab_Only_Template.xlsx`. In that case, the app estimates CEFR and Lexile from `Normal Ver.` before extracting Vocab & Click Words.

The download file is `Vocab_Click_Words_Analysis.xlsx`.

## Level Logic

- Input `Platform Level` is used only for generating `Easy Version` and `Difficult Version`.
- Stage 1 `Flagged Words` uses `Platform Level`: 1 -> A2, 2 -> B1, 3 -> B2, 4 -> C1. Words strictly above that ceiling are listed with up to three LCMS synonym-based suggestions.
- `Detected CEFR` and `Lexile` are estimated from `Base Text`.
- Vocab & Click Words filtering uses the estimated `Detected CEFR` with `lcms_cefr.csv`.
- Gemini suggestions are post-filtered with the LCMS DB. Known words below the detected CEFR are usually excluded, but a one-band-lower story-core click candidate can be kept when it is essential to comprehension.

Both stages support checkpoint JSON files for longer batches.
