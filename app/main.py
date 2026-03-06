from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse
import pandas as pd
import io
from typing import List

app = FastAPI(title="Excel összesítő")

# --- Mini frontend (root oldal) ---
@app.get("/", response_class=HTMLResponse)
def home():
    return """
<!doctype html>
<html lang="hu">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Excel összesítő</title>
  <style>
    :root { color-scheme: dark; }
    body {
      margin: 0;
      font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial;
      background: #0b1020;
      color: #e6e8ee;
    }
    .wrap { max-width: 820px; margin: 0 auto; padding: 28px 18px; }
    .card {
      background: rgba(255,255,255,0.06);
      border: 1px solid rgba(255,255,255,0.12);
      border-radius: 16px;
      padding: 18px;
      box-shadow: 0 12px 40px rgba(0,0,0,0.35);
    }
    h1 { font-size: 20px; margin: 0 0 10px; }
    p { margin: 8px 0; line-height: 1.5; color: rgba(230,232,238,0.9); }
    .row { display: flex; gap: 12px; flex-wrap: wrap; align-items: center; margin-top: 14px; }
    .file {
      flex: 1;
      min-width: 240px;
      padding: 10px 12px;
      background: rgba(0,0,0,0.25);
      border: 1px dashed rgba(255,255,255,0.18);
      border-radius: 12px;
    }
    input[type="file"] { width: 100%; }
    button {
      border: 0;
      border-radius: 12px;
      padding: 10px 14px;
      background: #6d5efc;
      color: white;
      font-weight: 700;
      cursor: pointer;
    }
    button:disabled { opacity: 0.6; cursor: not-allowed; }
    .small { font-size: 12px; opacity: 0.85; }
    .status { margin-top: 12px; padding: 10px 12px; border-radius: 12px; display:none; }
    .status.ok { display:block; background: rgba(45, 212, 191, 0.15); border: 1px solid rgba(45, 212, 191, 0.35); }
    .status.err { display:block; background: rgba(248, 113, 113, 0.12); border: 1px solid rgba(248, 113, 113, 0.35); }
    .spinner {
      width: 14px; height: 14px;
      border: 2px solid rgba(255,255,255,0.25);
      border-top-color: rgba(255,255,255,0.95);
      border-radius: 50%;
      display:inline-block;
      animation: spin 0.8s linear infinite;
      vertical-align: -2px;
      margin-right: 8px;
    }
    @keyframes spin { to { transform: rotate(360deg); } }
    .footer { margin-top: 10px; opacity: 0.7; font-size: 12px; }
    a { color: #bcb6ff; }
  </style>
</head>
<body>
  <div class="card">
    <h1>Excel összesítő</h1>
    <p>Tölts fel több <b>.xlsx</b> fájlt. A rendszer eldobja az A–E oszlopokat, majd cikkszám + név alapján összeadja a darabszámot.</p>

    <div class="row">
      <input id="files" type="file" multiple accept=".xlsx" />
      <button id="run" disabled>Összesítés</button>
    </div>

    <div class="status" id="status"></div>
    <div class="muted">
      Tipp: Ha üres az oldal, nézd meg a <a href="/docs">/docs</a> felületet is (API teszt).
    </div>
  </div>

<script>
  const filesEl = document.getElementById("files");
  const runBtn = document.getElementById("run");
  const statusEl = document.getElementById("status");

  filesEl.addEventListener("change", () => {
    runBtn.disabled = !(filesEl.files && filesEl.files.length);
    statusEl.textContent = filesEl.files.length ? (`Kiválasztva: ${filesEl.files.length} fájl`) : "";
  });

  runBtn.addEventListener("click", async () => {
    runBtn.disabled = true;
    statusEl.textContent = "Feldolgozás folyamatban…";

    const fd = new FormData();
    for (const f of filesEl.files) fd.append("files", f);

    try {
      const res = await fetch("/tools/excel/aggregate", { method: "POST", body: fd });
      if (!res.ok) {
        const msg = await res.text();
        statusEl.textContent = "Hiba: " + msg;
        runBtn.disabled = false;
        return;
      }

      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "osszesitett_struktura_tisztitott.xlsx";
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);

      statusEl.textContent = "Kész! A fájl letöltődött.";
    } catch (e) {
      statusEl.textContent = "Hálózati hiba: " + e;
    } finally {
      runBtn.disabled = false;
    }
  });
</script>
</body>
</html>
    """

def aggregate_excels(files: List[UploadFile]) -> bytes:
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

    combined_df = pd.concat(df_list, ignore_index=True)

    # A–E eldobása
    if combined_df.shape[1] <= 5:
        raise HTTPException(status_code=400, detail="Túl kevés oszlop: nem tudom eldobni az A–E oszlopokat.")
    combined_df = combined_df.iloc[:, 5:]

    original_columns = combined_df.columns.tolist()

    # várt pozíciók a vágás után: 0=cikkszám, 1=név, 6=darab
    if len(original_columns) <= 6:
        raise HTTPException(status_code=400, detail="Nem találom a darabszám oszlopot (várt index: 6 a vágás után).")

    cikkszam_col = combined_df.columns[0]
    nev_col = combined_df.columns[1]
    darab_col = combined_df.columns[6]

    agg_map = {col: "first" for col in combined_df.columns}
    agg_map[darab_col] = "sum"

    aggregated_df = (
        combined_df
        .groupby([cikkszam_col, nev_col], as_index=False)
        .agg(agg_map)
    )

    # oszlopsorrend vissza
    aggregated_df = aggregated_df[original_columns]

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
        headers={"Content-Disposition": 'attachment; filename="osszesitett_struktura_tisztitott.xlsx"'},
    )