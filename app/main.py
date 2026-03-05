from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
import pandas as pd
import io
from typing import List

app = FastAPI(title="AP Data Tools")

def aggregate_excels(files: List[UploadFile]) -> bytes:
    # Beolvasás memóriába
    df_list = []
    for f in files:
        if not f.filename.lower().endswith(".xlsx"):
            raise HTTPException(status_code=400, detail=f"Nem .xlsx fájl: {f.filename}")

        content = f.file.read()
        if not content:
            raise HTTPException(status_code=400, detail=f"Üres fájl: {f.filename}")

        try:
            df = pd.read_excel(io.BytesIO(content))
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"Hibás Excel ({f.filename}): {e}")

        df_list.append(df)

    if not df_list:
        raise HTTPException(status_code=400, detail="Nem érkezett feldolgozható fájl.")

    combined_df = pd.concat(df_list, ignore_index=True)

    # Eldobjuk az első 5 oszlopot (A–E)
    if combined_df.shape[1] <= 5:
        raise HTTPException(status_code=400, detail="Túl kevés oszlop: nem tudom eldobni az A–E oszlopokat.")

    combined_df = combined_df.iloc[:, 5:]

    # Oszlopsorrend megőrzése
    original_columns = combined_df.columns.tolist()

    # A levágás után:
    # cikkszám = 0 (eredetileg F)
    # név = 1 (eredetileg G)
    # darab = 6 (eredetileg L)
    if len(original_columns) <= 6:
        raise HTTPException(status_code=400, detail="Nem találom a darabszám oszlopot (várt index: 6 a vágás után).")

    cikkszam_col = combined_df.columns[0]
    nev_col = combined_df.columns[1]
    darab_col = combined_df.columns[6]

    # Aggregálás: darabszám sum, minden más first
    agg_map = {col: "first" for col in combined_df.columns}
    agg_map[darab_col] = "sum"

    aggregated_df = (
        combined_df
        .groupby([cikkszam_col, nev_col], as_index=False)
        .agg(agg_map)
    )

    # Oszlopsorrend vissza
    aggregated_df = aggregated_df[original_columns]

    # Excel export memóriába
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        aggregated_df.to_excel(writer, sheet_name="Osszesites", index=False)

    return out.getvalue()


@app.post("/tools/excel/aggregate")
async def excel_aggregate(files: List[UploadFile] = File(...)):
    if not files:
        raise HTTPException(status_code=400, detail="Nincs feltöltött fájl.")

    xlsx_bytes = aggregate_excels(files)

    return StreamingResponse(
        io.BytesIO(xlsx_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": 'attachment; filename="osszesitett_struktura_tisztitott.xlsx"'
        },
    )