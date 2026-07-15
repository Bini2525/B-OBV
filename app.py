"""
OBV Trend- & Divergenzanalyse
==============================

Streamlit-App zur automatisierten On-Balance-Volume (OBV) Analyse von Aktien.

Funktionen:
    - Ticker-Eingabe manuell (Textfeld, kommagetrennt) ODER per Excel-Upload
    - Automatische Auflösung europäischer Ticker ohne Börsen-Suffix
      (z. B. "BMW" -> "BMW.DE") über Suffix-Heuristik + Yahoo-Finance-Suche
    - Analyse-Zeitraum wählbar: 1 Woche, 1 Monat, 3 Monate, 6 Monate, 1 Jahr
    - Abruf historischer Kurs-/Volumendaten via yfinance
    - Berechnung des On-Balance-Volume (OBV)
    - Trend- und Divergenzanalyse (Kurs vs. OBV) über die letzten Handelstage
    - Durchschnittliches Analystenkursziel je Ticker (sofern von Yahoo Finance geführt)
    - Interaktive Ergebnistabelle in der App (deutsches Zahlenformat)
    - Download des Ergebnisses als formatierte Excel-Datei (.xlsx)

Hinweis: Diese App dient ausschließlich Informationszwecken und stellt
keine Anlageberatung dar.

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
# "1 Woche" wird über "5d" (5 Handelstage) abgebildet, da yfinance keinen
# eigenen 1-Wochen-Period-String kennt.
PERIOD_OPTIONS: dict[str, str] = {
    "1 Woche": "5d",
    "1 Monat": "1mo",
    "3 Monate": "3mo",
    "6 Monate": "6mo",
    "1 Jahr": "1y",
}
DEFAULT_PERIOD_LABEL = "6 Monate"

# Mindestanzahl an Handelstagen, ab der die Trend-/Divergenzanalyse als
# statistisch belastbar gilt. Bei kürzeren Zeiträumen (z. B. "1 Woche")
# wird das Ergebnis trotzdem berechnet, aber als eingeschränkt gekennzeichnet.
MIN_RELIABLE_DAYS = 10

# Anzahl der Handelstage, die für die Trend-/Divergenzanalyse betrachtet werden.
LOOKBACK_DAYS = 15  # liegt im geforderten Korridor von 10 bis 20 Handelstagen

# Schwellenwert für die Trendklassifikation (relative Steigung pro Tag).
# Ein Wert von 0.001 entspricht einer durchschnittlichen Veränderung von
# ca. 0,1 % pro Handelstag und dient dazu, "Rauschen" von einem echten
# Trend zu unterscheiden.
TREND_SLOPE_THRESHOLD = 0.001

# Spaltennamen, nach denen beim Excel-Import gesucht wird (case-insensitive).
TICKER_COLUMN_CANDIDATES = ["ticker", "symbol", "aktie", "wertpapier"]

# Pflicht-Disclaimer für alle Aktien-/Finanzanalyse-Apps.
DISCLAIMER_TEXT = (
    "Hinweis: Diese App dient ausschließlich Informationszwecken und stellt "
    "keine Anlageberatung dar. Keine Kauf- oder Verkaufsempfehlung."
)

# Gängige Börsen-Suffixe für Yahoo Finance, nach denen gesucht wird, wenn ein
# roher Ticker (z. B. "BMW" statt "BMW.DE") nicht direkt gefunden wird.
# Schwerpunkt DACH/Westeuropa, da für europäische Broker-Exports typisch.
EXCHANGE_SUFFIXES = [
    ".DE",  # Xetra (Deutschland)
    ".F",   # Frankfurt (Alternative zu Xetra)
    ".SW",  # SIX Swiss Exchange
    ".L",   # London Stock Exchange
    ".AS",  # Euronext Amsterdam
    ".PA",  # Euronext Paris
    ".MI",  # Borsa Italiana Mailand
    ".MC",  # Bolsa de Madrid
    ".BR",  # Euronext Brüssel
    ".VI",  # Wiener Börse
]


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
# Ticker-Auflösung (Börsen-Suffix-Heuristik + Yahoo-Suche) & Datenabruf
# ---------------------------------------------------------------------------

def _try_symbol(symbol: str, period: str):
    """
    Versucht, für ein konkretes Yahoo-Finance-Symbol Kursdaten zu laden.

    Gibt (ticker_obj, hist) zurück, wenn verwertbare Daten vorliegen,
    sonst None. Wirft absichtlich keine Exception - wird als "Testballon"
    innerhalb der Auflösungs-Kette verwendet.
    """
    try:
        ticker_obj = yf.Ticker(symbol)
        hist = ticker_obj.history(period=period, auto_adjust=False)
        if hist is None or hist.empty:
            return None
        if "Close" not in hist.columns or "Volume" not in hist.columns:
            return None
        hist = hist[["Close", "Volume"]].dropna()
        if hist.empty:
            return None
        return ticker_obj, hist
    except Exception:
        return None


def resolve_ticker(raw_ticker: str, period: str) -> tuple[str, "yf.Ticker", pd.DataFrame]:
    """
    Löst einen rohen Ticker-Code in ein gültiges Yahoo-Finance-Symbol auf und
    lädt dabei direkt die passende Kurshistorie.

    Viele europäische Broker-Exports (z. B. Trade Republic, comdirect) führen
    Aktien ohne Börsen-Suffix (z. B. "BMW" statt "BMW.DE"). Yahoo Finance
    benötigt für Nicht-US-Börsen aber i. d. R. ein Suffix. Auflösungs-Reihenfolge:

        1. Ticker wie eingegeben probieren (deckt US-Ticker sowie bereits
           korrekt angegebene Symbole wie "BMW.DE" ab).
        2. Yahoo-Finance-Volltextsuche (yfinance.Search) - findet in der
           Praxis meist direkt das richtige Symbol samt Börsenplatz.
        3. Als letzter Fallback: gängige europäische Börsensuffixe (siehe
           EXCHANGE_SUFFIXES) systematisch durchprobieren.

    Gibt (aufgelöstes_symbol, ticker_obj, hist) zurück oder wirft eine
    ValueError, wenn keiner der drei Wege zu verwertbaren Daten führt.
    """
    # 1) Roher Ticker wie eingegeben.
    found = _try_symbol(raw_ticker, period)
    if found is not None:
        return raw_ticker, found[0], found[1]

    # 2) Yahoo-Finance-Suche (liefert i. d. R. bereits das korrekte,
    #    börsenspezifische Symbol als Top-Treffer).
    try:
        search_results = yf.Search(raw_ticker, max_results=5).quotes
    except Exception:
        search_results = []

    for quote in search_results:
        symbol = quote.get("symbol")
        if not symbol:
            continue
        found = _try_symbol(symbol, period)
        if found is not None:
            return symbol, found[0], found[1]

    # 3) Suffix-Heuristik als letzter Fallback (nur wenn kein Suffix bereits
    #    im Ticker enthalten ist - sonst würden unsinnige Kombinationen wie
    #    "BMW.DE.SW" entstehen).
    if "." not in raw_ticker:
        for suffix in EXCHANGE_SUFFIXES:
            candidate = f"{raw_ticker}{suffix}"
            found = _try_symbol(candidate, period)
            if found is not None:
                return candidate, found[0], found[1]

    raise ValueError(
        f"Kein gültiges Yahoo-Finance-Symbol für '{raw_ticker}' gefunden - "
        f"auch nicht über die Yahoo-Suche oder gängige Börsensuffixe "
        f"(.DE/.SW/.L/...). Ticker ggf. mit explizitem Börsenkürzel angeben, "
        f"z. B. '{raw_ticker}.DE'."
    )


def fetch_price_history(ticker: str, period: str) -> tuple[pd.DataFrame, dict]:
    """
    Ruft historische Kurs- und Volumendaten für einen Ticker ab (inkl.
    automatischer Ticker-Auflösung, siehe resolve_ticker).

    Gibt ein DataFrame mit den Spalten 'Close' und 'Volume' sowie ein
    Meta-Dict (Name, Handelswährung, durchschnittliches Analystenkursziel,
    aufgelöstes Yahoo-Symbol) zurück. Wirft eine Exception, wenn keine
    verwertbaren Daten vorliegen (wird vom Aufrufer abgefangen).
    """
    resolved_symbol, ticker_obj, hist = resolve_ticker(ticker, period)

    meta = {
        "currency": "n/a",
        "avg_analyst_target": None,
        "name": resolved_symbol,
        "resolved_symbol": resolved_symbol,
    }

    # Firmenname, Handelswährung & durchschnittliches Analystenkursziel
    # ermitteln. Alles nicht kritisch für die OBV-Analyse selbst (nicht jeder
    # Ticker hat z. B. Analysten-Coverage), daher robust mit Fallback statt
    # Abbruch.
    try:
        info = ticker_obj.info
        if info:
            meta["currency"] = info.get("currency") or meta["currency"]
            target = info.get("targetMeanPrice")
            if target is not None:
                meta["avg_analyst_target"] = float(target)
            name = info.get("longName") or info.get("shortName")
            if name:
                meta["name"] = name
    except Exception:
        pass

    if meta["currency"] == "n/a":
        try:
            fast_info = getattr(ticker_obj, "fast_info", None)
            if fast_info is not None:
                meta["currency"] = fast_info.get("currency") or "n/a"
        except Exception:
            pass

    return hist, meta


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
        bewertung = "Aufwärtstrend bestätigt"
    elif price_trend == "down" and obv_trend == "down":
        bewertung = "Abwärtstrend bestätigt"
    elif price_trend in ("down", "flat") and obv_trend == "up":
        bewertung = "Bullische Divergenz (Kaufsignal)"
    elif price_trend == "up" and obv_trend in ("down", "flat"):
        bewertung = "Bärische Divergenz (Verkaufssignal)"
    else:
        bewertung = "Keine klare Richtung"

    # Bei kurzen Zeiträumen (z. B. "1 Woche") steht weniger als
    # MIN_RELIABLE_DAYS Handelstage zur Verfügung - Ergebnis wird trotzdem
    # ausgegeben, aber transparent als weniger belastbar gekennzeichnet.
    if len(window) < MIN_RELIABLE_DAYS:
        bewertung += " (eingeschränkte Datenbasis)"

    return bewertung


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
        hist, meta = fetch_price_history(ticker, period)
        hist_with_obv = calculate_obv(hist)

        if len(hist_with_obv) < 2:
            return None, f"{ticker}: Zu wenige Datenpunkte für eine Analyse."

        bewertung = analyze_trend_divergence(hist_with_obv, LOOKBACK_DAYS)

        avg_target = meta.get("avg_analyst_target")

        result = {
            "Ticker": ticker,
            "Name": meta.get("name") or ticker,
            "Währung": meta.get("currency", "n/a"),
            "Aktueller Kurs": round(float(hist_with_obv["Close"].iloc[-1]), 2),
            "Durchschnittliches Analystenkursziel": round(avg_target, 2) if avg_target is not None else None,
            "Letztes Volumen": int(hist_with_obv["Volume"].iloc[-1]),
            "Aktueller OBV-Wert": int(hist_with_obv["OBV"].iloc[-1]),
            "Trend-Bestätigung / Divergenz": bewertung,
        }
        return result, None

    except Exception as exc:  # noqa: BLE001 - bewusst breit, damit kein Ticker die App killt
        return None, f"{ticker}: Analyse fehlgeschlagen ({exc})."


# ---------------------------------------------------------------------------
# Deutsches Zahlenformat für die Bildschirmanzeige
# ---------------------------------------------------------------------------

def format_de_number(value, decimals: int = 2) -> str:
    """
    Formatiert eine Zahl im deutschen Format (Tausenderpunkt, Komma als
    Dezimaltrennzeichen), z. B. 1234.5 -> "1.234,50".

    Gibt einen leeren String zurück, wenn kein Wert vorhanden ist
    (None/NaN) - relevant z. B. für fehlendes Analystenkursziel.
    """
    if value is None:
        return ""
    try:
        if isinstance(value, float) and np.isnan(value):
            return ""
    except TypeError:
        pass

    # Trick: zunächst im US-Format formatieren (Komma=Tausender, Punkt=Dezimal),
    # anschließend Trennzeichen tauschen -> deutsches Format.
    us_formatted = f"{value:,.{decimals}f}"
    return us_formatted.replace(",", "§").replace(".", ",").replace("§", ".")


def build_display_dataframe(result_df: pd.DataFrame) -> pd.DataFrame:
    """
    Erstellt eine Anzeige-Kopie des Ergebnis-DataFrames mit Zahlen im
    deutschen Format - ausschließlich für die Bildschirmanzeige in
    st.dataframe(). Der Excel-Export (build_excel_bytes) arbeitet weiterhin
    mit den numerischen Rohwerten aus result_df, damit Excel eigene
    Zahlenformate korrekt anwenden kann.
    """
    display_df = result_df.copy()

    two_decimal_cols = ["Aktueller Kurs", "Durchschnittliches Analystenkursziel"]
    zero_decimal_cols = ["Letztes Volumen", "Aktueller OBV-Wert"]

    for col in two_decimal_cols:
        if col in display_df.columns:
            display_df[col] = display_df[col].apply(lambda v: format_de_number(v, 2))
    for col in zero_decimal_cols:
        if col in display_df.columns:
            display_df[col] = display_df[col].apply(lambda v: format_de_number(v, 0))

    return display_df


def _max_text_length(values, minimum: int = 0) -> int:
    """
    Ermittelt die maximale Textlänge einer Spalte, robust gegenüber None/NaN
    und unabhängig vom pandas-Backend (numpy-object vs. Arrow-backed).

    Hintergrund: `series.astype(str).map(len)` ist NICHT sicher, wenn die
    Spalte None-Werte enthält - je nach pandas-/pyarrow-Version bleibt der
    Null-Wert nach astype(str) ein NA-Sentinel statt der Textstring "None"
    zu werden, wodurch len() mit TypeError abbricht (genau dieser Fehler
    trat in Produktion bei fehlendem Analystenkursziel auf). Deshalb hier
    bewusst eine explizite pd.notna()-Prüfung statt astype(str) + map(len).
    """
    lengths = [len(str(v)) for v in values if pd.notna(v)]
    return max(lengths, default=minimum)


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

        # Textspalten: linksbündig. Alle übrigen Spalten (Zahlen) rechtsbündig,
        # 2 Nachkommastellen + Tausenderpunkt bzw. 0 Nachkommastellen gemäß
        # globalem Zahlenformat-Standard.
        text_columns = {"Ticker", "Name", "Währung", "Trend-Bestätigung / Divergenz"}
        number_format_map = {
            "Aktueller Kurs": "#,##0.00",
            "Durchschnittliches Analystenkursziel": "#,##0.00",
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

            # Spaltenbreite: Mindestbreite 13, sonst am längsten Inhalt
            # ausgerichtet, damit keine "####"-Anzeige entsteht. None/NaN
            # werden dabei übersprungen statt einen Fehler zu werfen (siehe
            # _max_text_length).
            max_content_len = max(
                len(str(col_name)),
                _max_text_length(result_df[col_name].tolist()) if n_rows > 0 else 0,
            )
            worksheet.column_dimensions[col_letter].width = max(13, max_content_len + 2)

        worksheet.freeze_panes = "A2"

        # Disclaimer unterhalb der Tabelle: keine Anlageberatung.
        disclaimer_row = n_rows + 3
        worksheet.cell(row=disclaimer_row, column=1, value=DISCLAIMER_TEXT)
        worksheet.cell(row=disclaimer_row, column=1).font = Font(name="Arial", size=9, italic=True)

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
        "Handelswährung des jeweiligen Tickers angezeigt (siehe Spalte 'Währung'). "
        "Ticker ohne Börsenkürzel (z. B. 'BMW' statt 'BMW.DE') werden automatisch "
        "über gängige europäische Börsensuffixe und die Yahoo-Finance-Suche aufgelöst."
    )
    st.warning(f"⚠️ {DISCLAIMER_TEXT}")

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
    st.dataframe(build_display_dataframe(result_df), use_container_width=True, hide_index=True)

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
