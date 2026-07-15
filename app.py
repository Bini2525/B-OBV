"""
OBV Trend- & Divergenzanalyse
==============================

Streamlit-App zur automatisierten On-Balance-Volume (OBV) Analyse von Aktien.

Funktionen:
    - Ticker-Eingabe manuell (Textfeld, kommagetrennt) ODER per Excel-Upload
    - Abruf historischer Kurs-/Volumendaten via yfinance
    - Berechnung des On-Balance-Volume (OBV)
    - Trend- und Divergenzanalyse (Kurs vs. OBV) über die letzten Handelstage
    - Interaktive Ergebnistabelle in der App
    - Download des Ergebnisses als formatierte Excel-Datei (.xlsx)

Hosting:
    Kann unverändert über GitHub + Streamlit Community Cloud deployed werden.
    Einstiegsdatei: app.py | Abhängigkeiten: requirements.txt
"""

from __future__ import annotations

import io
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st
import yfinance as yf
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Konstanten / Konfiguration
# ---------------------------------------------------------------------------

# Zuordnung der UI-Auswahl auf die von yfinance erwarteten Period-Strings.
PERIOD_OPTIONS: dict[str, str] = {
    "3 Monate": "3mo",
    "6 Monate": "6mo",
    "1 Jahr": "1y",
}
DEFAULT_PERIOD_LABEL = "6 Monate"

# Anzahl der Handelstage, die für die Trend-/Divergenzanalyse betrachtet werden.
LOOKBACK_DAYS = 15  # liegt im geforderten Korridor von 10 bis 20 Handelstagen

# Schwellenwert für die Trendklassifikation (relative Steigung pro Tag).
# Ein Wert von 0.001 entspricht einer durchschnittlichen Veränderung von
# ca. 0,1 % pro Handelstag und dient dazu, "Rauschen" von einem echten
# Trend zu unterscheiden.
TREND_SLOPE_THRESHOLD = 0.001

# Spaltennamen, nach denen beim Excel-Import gesucht wird (case-insensitive).
TICKER_COLUMN_CANDIDATES = ["ticker", "symbol", "aktie", "wertpapier"]


# ---------------------------------------------------------------------------
# Hilfsfunktionen: Ticker-Eingabe (manuell & Excel-Upload)
# ---------------------------------------------------------------------------

def parse_manual_tickers(raw_text: str) -> list[str]:
    """Wandelt eine kommagetrennte Texteingabe in eine bereinigte Ticker-Liste um."""
    if not raw_text:
        return []
    tickers = [t.strip().upper() for t in raw_text.split(",")]
    return [t for t in tickers if t]  # leere Einträge entfernen


def extract_tickers_from_excel(uploaded_file) -> list[str]:
    """
    Liest eine hochgeladene Excel-Datei ein und extrahiert die Ticker-Spalte.

    Sucht case-insensitive nach gängigen Spaltennamen (ticker, symbol, aktie,
    wertpapier). Wird keine passende Spalte gefunden, wird die erste Spalte
    der Tabelle verwendet.
    """
    df = pd.read_excel(uploaded_file)

    if df.empty or df.shape[1] == 0:
        return []

    # Case-insensitive Suche nach einer passenden Spaltenbezeichnung.
    ticker_column = None
    for col in df.columns:
        if str(col).strip().lower() in TICKER_COLUMN_CANDIDATES:
            ticker_column = col
            break

    # Fallback: erste Spalte der Tabelle verwenden.
    if ticker_column is None:
        ticker_column = df.columns[0]

    raw_values = df[ticker_column].dropna().astype(str).tolist()
    tickers = [v.strip().upper() for v in raw_values]
    return [t for t in tickers if t]


def get_ticker_list(input_mode: str, manual_text: str, uploaded_file) -> list[str]:
    """Ermittelt je nach gewähltem Eingabemodus die finale, deduplizierte Ticker-Liste."""
    if input_mode == "Manuelle Eingabe":
        tickers = parse_manual_tickers(manual_text)
    else:
        tickers = extract_tickers_from_excel(uploaded_file)

    # Reihenfolge beibehalten, aber Duplikate entfernen.
    seen = set()
    unique_tickers = []
    for t in tickers:
        if t not in seen:
            seen.add(t)
            unique_tickers.append(t)
    return unique_tickers


# ---------------------------------------------------------------------------
# Datenabruf & OBV-Berechnung
# ---------------------------------------------------------------------------

def fetch_price_history(ticker: str, period: str) -> tuple[pd.DataFrame, str]:
    """
    Ruft historische Kurs- und Volumendaten für einen Ticker ab.

    Gibt ein DataFrame mit den Spalten 'Close' und 'Volume' sowie die
    gemeldete Handelswährung zurück. Wirft eine Exception, wenn keine
    verwertbaren Daten vorliegen (wird vom Aufrufer abgefangen).
    """
    ticker_obj = yf.Ticker(ticker)
    hist = ticker_obj.history(period=period, auto_adjust=False)

    if hist is None or hist.empty:
        raise ValueError(f"Keine Kursdaten für '{ticker}' gefunden.")

    if "Close" not in hist.columns or "Volume" not in hist.columns:
        raise ValueError(f"Unvollständige Daten für '{ticker}' (Close/Volume fehlt).")

    hist = hist[["Close", "Volume"]].dropna()
    if hist.empty:
        raise ValueError(f"Keine verwertbaren Kurs-/Volumendaten für '{ticker}'.")

    # Handelswährung ermitteln (nicht kritisch, daher robust mit Fallback).
    currency = "n/a"
    try:
        fast_info = getattr(ticker_obj, "fast_info", None)
        if fast_info is not None:
            currency = fast_info.get("currency") or "n/a"
    except Exception:
        currency = "n/a"

    return hist, currency


def calculate_obv(hist: pd.DataFrame) -> pd.DataFrame:
    """
    Berechnet das On-Balance-Volume (OBV) für die übergebene Kurshistorie.

    Regeln:
        Close(t) > Close(t-1)  -> OBV(t) = OBV(t-1) + Volume(t)
        Close(t) < Close(t-1)  -> OBV(t) = OBV(t-1) - Volume(t)
        Close(t) == Close(t-1) -> OBV(t) = OBV(t-1)
    """
    df = hist.copy()
    close_diff = df["Close"].diff()

    # Richtung: +1 (Anstieg), -1 (Rückgang), 0 (unverändert).
    direction = np.sign(close_diff).fillna(0)

    # Signierte Volumina aufkumulieren -> OBV. Der erste Wert startet bei 0.
    signed_volume = direction * df["Volume"]
    df["OBV"] = signed_volume.cumsum()
    df["OBV"] = df["OBV"].fillna(0)

    return df


# ---------------------------------------------------------------------------
# Trend- & Divergenzanalyse
# ---------------------------------------------------------------------------

def classify_trend(series: pd.Series, threshold: float = TREND_SLOPE_THRESHOLD) -> str:
    """
    Klassifiziert den Trend einer Zeitreihe als 'up', 'down' oder 'flat'.

    Methode: lineare Regression (Steigung) über das Analysefenster,
    normalisiert auf das mittlere Niveau der Reihe, damit die Klassifikation
    unabhängig von der absoluten Größenordnung (Kurs vs. OBV) funktioniert.
    """
    series = series.dropna()
    if len(series) < 2:
        return "flat"

    x = np.arange(len(series))
    slope = np.polyfit(x, series.values, 1)[0]

    avg_level = np.mean(np.abs(series.values))
    if avg_level == 0:
        avg_level = 1.0  # Division durch 0 vermeiden

    normalized_slope = slope / avg_level

    if normalized_slope > threshold:
        return "up"
    if normalized_slope < -threshold:
        return "down"
    return "flat"


def analyze_trend_divergence(df: pd.DataFrame, lookback: int = LOOKBACK_DAYS) -> str:
    """
    Vergleicht den Kurstrend mit dem OBV-Trend über die letzten `lookback`
    Handelstage und leitet daraus eine Bewertung ab.

    Regeln:
        Kurs up   + OBV up            -> Aufwärtstrend bestätigt
        Kurs down + OBV down          -> Abwärtstrend bestätigt
        Kurs down/flat + OBV up       -> Bullische Divergenz (Kaufsignal)
        Kurs up + OBV down/flat       -> Bärische Divergenz (Verkaufssignal)
        alle anderen Kombinationen    -> Keine klare Richtung
    """
    window = df.tail(lookback)

    price_trend = classify_trend(window["Close"])
    obv_trend = classify_trend(window["OBV"])

    if price_trend == "up" and obv_trend == "up":
        return "Aufwärtstrend bestätigt"
    if price_trend == "down" and obv_trend == "down":
        return "Abwärtstrend bestätigt"
    if price_trend in ("down", "flat") and obv_trend == "up":
        return "Bullische Divergenz (Kaufsignal)"
    if price_trend == "up" and obv_trend in ("down", "flat"):
        return "Bärische Divergenz (Verkaufssignal)"
    return "Keine klare Richtung"


# ---------------------------------------------------------------------------
# Analyse-Pipeline pro Ticker
# ---------------------------------------------------------------------------

def analyze_ticker(ticker: str, period: str) -> tuple[dict | None, str | None]:
    """
    Führt die vollständige Analyse für einen einzelnen Ticker aus.

    Gibt ein Tupel (Ergebnis-Dict, Fehlermeldung) zurück. Genau eines der
    beiden Elemente ist None - so kann der Aufrufer sauber zwischen Erfolg
    und Fehlschlag unterscheiden, ohne die App abstürzen zu lassen.
    """
    try:
        hist, currency = fetch_price_history(ticker, period)
        hist_with_obv = calculate_obv(hist)

        if len(hist_with_obv) < 2:
            return None, f"{ticker}: Zu wenige Datenpunkte für eine Analyse."

        bewertung = analyze_trend_divergence(hist_with_obv, LOOKBACK_DAYS)

        result = {
            "Ticker": ticker,
            "Währung": currency,
            "Aktueller Kurs": round(float(hist_with_obv["Close"].iloc[-1]), 2),
            "Letztes Volumen": int(hist_with_obv["Volume"].iloc[-1]),
            "Aktueller OBV-Wert": int(hist_with_obv["OBV"].iloc[-1]),
            "Trend-Bestätigung / Divergenz": bewertung,
        }
        return result, None

    except Exception as exc:  # noqa: BLE001 - bewusst breit, damit kein Ticker die App killt
        return None, f"{ticker}: Analyse fehlgeschlagen ({exc})."


# ---------------------------------------------------------------------------
# Excel-Export mit Formatierung
# ---------------------------------------------------------------------------

def build_excel_bytes(result_df: pd.DataFrame) -> bytes:
    """Erstellt eine formatierte .xlsx-Datei (Bytes) aus dem Ergebnis-DataFrame."""
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        result_df.to_excel(writer, index=False, sheet_name="OBV-Analyse")
        worksheet = writer.sheets["OBV-Analyse"]

        header_font = Font(name="Arial", size=11, bold=True)
        data_font = Font(name="Arial", size=11)
        data_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        thin_black = Side(border_style="thin", color="000000")
        table_border = Border(left=thin_black, right=thin_black, top=thin_black, bottom=thin_black)

        text_columns = {"Ticker", "Währung", "Trend-Bestätigung / Divergenz"}
        number_format_map = {
            "Aktueller Kurs": "#,##0.00",
            "Letztes Volumen": "#,##0",
            "Aktueller OBV-Wert": "#,##0",
        }

        n_rows, n_cols = result_df.shape

        for col_idx, col_name in enumerate(result_df.columns, start=1):
            col_letter = get_column_letter(col_idx)

            # Kopfzeile: fett, kein Zeilenumbruch, Arial 11.
            header_cell = worksheet.cell(row=1, column=col_idx)
            header_cell.font = header_font
            header_cell.alignment = Alignment(
                horizontal="left" if col_name in text_columns else "right",
                vertical="top",
                wrap_text=False,
            )
            header_cell.border = table_border

            # Datenzeilen: berechnete/nicht editierbare Werte -> blau hinterlegt.
            for row_idx in range(2, n_rows + 2):
                cell = worksheet.cell(row=row_idx, column=col_idx)
                cell.font = data_font
                cell.fill = data_fill
                cell.border = table_border
                cell.alignment = Alignment(
                    horizontal="left" if col_name in text_columns else "right",
                    vertical="top",
                    wrap_text=True,
                )
                if col_name in number_format_map:
                    cell.number_format = number_format_map[col_name]

            # Spaltenbreite: Mindestbreite 13, sonst am längsten Inhalt ausgerichtet,
            # damit keine "####"-Anzeige entsteht.
            max_content_len = len(str(col_name))
            if n_rows > 0:
                col_values = result_df[col_name].astype(str)
                max_content_len = max(max_content_len, col_values.map(len).max())
            worksheet.column_dimensions[col_letter].width = max(13, max_content_len + 2)

        worksheet.freeze_panes = "A2"

    buffer.seek(0)
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------------------------

def main() -> None:
    st.set_page_config(page_title="OBV Trend- & Divergenzanalyse", layout="wide")
    st.title("📊 OBV Trend- & Divergenzanalyse")
    st.caption(
        "Berechnet das On-Balance-Volume (OBV) für beliebige Aktien und vergleicht "
        "den Kurstrend mit dem OBV-Trend, um Trendbestätigungen und Divergenzen "
        "zu erkennen. Kurse werden in der von Yahoo Finance gemeldeten "
        "Handelswährung des jeweiligen Tickers angezeigt (siehe Spalte 'Währung')."
    )

    # --- Sidebar: Eingabe -------------------------------------------------
    st.sidebar.header("Eingabe")

    input_mode = st.sidebar.radio(
        "Wie möchtest du die Ticker eingeben?",
        options=["Manuelle Eingabe", "Excel-Datei hochladen"],
    )

    manual_text = ""
    uploaded_file = None

    if input_mode == "Manuelle Eingabe":
        manual_text = st.sidebar.text_input(
            "Ticker (einzeln oder kommagetrennt)",
            placeholder="z. B. AAPL oder AAPL, MSFT, TSLA",
        )
    else:
        uploaded_file = st.sidebar.file_uploader(
            "Excel-Datei mit Ticker-Liste (.xlsx / .xls)",
            type=["xlsx", "xls"],
            help=(
                "Es wird automatisch nach einer Spalte 'Ticker', 'Symbol', 'Aktie' "
                "oder 'Wertpapier' gesucht. Wird keine gefunden, wird die erste "
                "Spalte der Tabelle verwendet."
            ),
        )

    period_label = st.sidebar.selectbox(
        "Analyse-Zeitraum",
        options=list(PERIOD_OPTIONS.keys()),
        index=list(PERIOD_OPTIONS.keys()).index(DEFAULT_PERIOD_LABEL),
    )

    start_analysis = st.sidebar.button("Analyse starten", type="primary")

    # --- Hauptbereich: Analyse & Ausgabe -----------------------------------
    if not start_analysis:
        st.info("Ticker eingeben bzw. Excel-Datei hochladen, Zeitraum wählen und auf "
                 "'Analyse starten' klicken.")
        return

    tickers = get_ticker_list(input_mode, manual_text, uploaded_file)

    if not tickers:
        st.warning(
            "Es wurden keine Ticker gefunden. Bitte prüfe deine Eingabe bzw. die "
            "hochgeladene Excel-Datei."
        )
        return

    period = PERIOD_OPTIONS[period_label]

    results: list[dict] = []
    warnings: list[str] = []

    progress_bar = st.progress(0.0, text="Analyse läuft ...")
    for i, ticker in enumerate(tickers, start=1):
        result, error = analyze_ticker(ticker, period)
        if result is not None:
            results.append(result)
        if error is not None:
            warnings.append(error)
        progress_bar.progress(i / len(tickers), text=f"Analysiere {ticker} ...")
    progress_bar.empty()

    # Fehlerhafte / delistete Ticker freundlich melden, App läuft dennoch weiter.
    if warnings:
        st.warning(
            "Folgende Ticker konnten nicht analysiert werden und wurden "
            "übersprungen:\n\n" + "\n".join(f"- {w}" for w in warnings)
        )

    if not results:
        st.error("Für keinen der eingegebenen Ticker konnten Daten ermittelt werden.")
        return

    result_df = pd.DataFrame(results)

    st.subheader("Ergebnis")
    st.dataframe(result_df, use_container_width=True, hide_index=True)

    # --- Excel-Export -------------------------------------------------------
    excel_bytes = build_excel_bytes(result_df)
    file_date = datetime.today().strftime("%d.%m.%y")

    st.download_button(
        label="📥 Ergebnis als Excel herunterladen",
        data=excel_bytes,
        file_name=f"{file_date}_OBV_Analyse.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()
