# Story Info + Vocab & Click Words

Separate Streamlit app that extends the existing Vocab & Click Words workflow.

## Input

Download `Story_Confirmed_Template.xlsx` in the app and fill:

- `ID`
- `Title`
- `Level`
- `Base Text`

`Level` accepts `1-4` or `Pre A1`, `A1`, `A2`, `B1`, `B2`, `C1`, `C2`.

## Level Logic

- Input `Level` is used only for generating `Easy Version` and `Difficult Version`.
- `Detected CEFR` and `Lexile` are estimated from `Base Text`.
- Vocab & Click Words filtering uses the estimated `Detected CEFR` with `lcms_cefr.csv`.

## Output

The app downloads `Story_Info_Vocab_Analysis.xlsx` with:

- `Combined`
- `Story_Info`
- `Vocab_Click_Words`

It also supports a checkpoint JSON to resume repeated runs without reprocessing completed IDs.
