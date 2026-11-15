# CineScope - movie analytics website (Streamlit 1.46 or newer)
# Run with:  streamlit run app.py
# Put khushi data (1).xlsx in the SAME folder as this file. It loads automatically.

import glob
import html
import io
import os
import re

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="CineScope | Movies", page_icon="🎬", layout="wide",
                   initial_sidebar_state="expanded")

GOLD, TEAL, ROSE, SKY = "#FFC83D", "#3CE0CD", "#FF6B6B", "#7AA2FF"
SEQ = [GOLD, TEAL, ROSE, SKY, "#C49BFF", "#9BE564", "#FF9F68", "#5EE6EB", "#FF8AD8", "#B8C1FF"]
WHITE = "#FFFFFF"
PAGE_SIZE = 12

# Optional: if the file is not found automatically, paste its full path here, for example
# DATA_PATH = r"C:\Users\DELL\Downloads\CineScope\khushi data.xlsx"
DATA_PATH = ""

# ------------------------------------------------------------------
# THEME  (dark background, bright white text)
# ------------------------------------------------------------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Outfit:wght@400;500;600;700;800&family=Fraunces:opsz,wght@9..144,700;9..144,800&display=swap');
html, body, [class*="css"] { font-family: 'Outfit', sans-serif; }
#MainMenu, footer, [data-testid="stAppDeployButton"], [data-testid="stDecoration"] { display: none !important; }
header[data-testid="stHeader"] { background: rgba(10,11,16,.92); backdrop-filter: blur(8px); border-bottom: 1px solid #2B2F3B; }
.block-container { padding-top: 4.5rem; max-width: 1320px; }
.stApp { background: radial-gradient(1100px 500px at 85% -10%, rgba(255,200,61,.14), transparent 60%),
         radial-gradient(900px 500px at -5% 10%, rgba(60,224,205,.10), transparent 55%), #0A0B10; color: #FFFFFF; }
[data-testid="stSidebar"] { background: #0F1118; border-right: 1px solid #2B2F3B; }

/* bright white text everywhere */
.stApp :is(p, label, li, h1, h2, h3, h4, h5, h6, td, th) { color: #FFFFFF; }
.stApp h1, .stApp h2, .stApp h3, .stApp h4 { color: #FFFFFF !important; letter-spacing: -.01em; font-weight: 700; }
[data-testid="stWidgetLabel"] p, [data-testid="stMetricLabel"] p, [data-testid="stMetricValue"],
.stTabs [data-baseweb="tab"] p, [data-testid="stCaptionContainer"] p, [data-testid="stCaptionContainer"] { color: #FFFFFF !important; }
[data-testid="stCaptionContainer"] p { opacity: .92; }
input, textarea, [data-baseweb="select"] * { color: #FFFFFF !important; }
.stTabs [data-baseweb="tab"] { font-weight: 600; }
.stTabs [aria-selected="true"] p { color: #FFC83D !important; }

/* hero */
.hero { position: relative; overflow: hidden; padding: 64px 48px 56px; border-radius: 28px; border: 1px solid #2F3342;
        margin-bottom: 14px; background: linear-gradient(120deg, #141722 0%, #1A1913 55%, #2A1F0B 100%); }
.hero:before { content: ""; position: absolute; left: 0; right: 0; top: 0; height: 14px; opacity: .35;
        background: repeating-linear-gradient(90deg, #FFC83D 0 18px, transparent 18px 34px); }
.hero:after { content: "🎬"; position: absolute; right: 48px; bottom: 10px; font-size: 170px; opacity: .08; }
.stApp .eyebrow { color: #FFC83D; font-size: 13px; font-weight: 700; letter-spacing: .18em; }
.hero-title { font-family: 'Fraunces', serif; font-size: 58px; font-weight: 800; line-height: 1.05; margin-top: 14px; max-width: 800px; color: #FFFFFF; }
.stApp .hero-title span { color: #FFC83D; }
.stApp .hero-sub { color: #EDEEF4; font-size: 18px; margin-top: 18px; max-width: 640px; line-height: 1.5; }
.chips { margin-top: 18px; display: flex; gap: 8px; flex-wrap: wrap; }
.stApp .chip { padding: 5px 13px; border: 1px solid #5A5E70; border-radius: 999px; font-size: 13px; color: #FFFFFF; }

.st-key-cta [data-testid="stPageLink-NavLink"] { border: 1px solid #FFC83D; border-radius: 12px; justify-content: center;
        padding: .6rem 1rem; background: rgba(255,200,61,.08); }
.st-key-cta [data-testid="stPageLink-NavLink"]:hover { background: #FFC83D; }
.st-key-cta [data-testid="stPageLink-NavLink"]:hover p { color: #15120A !important; }
.st-key-cta p { font-weight: 700; color: #FFC83D !important; }
.st-key-spot { background: linear-gradient(120deg, #151822, #1D1A12); border: 1px solid #2F3342; border-radius: 22px; padding: 26px; }

/* cards */
.metric-card { background: #12141C; border: 1px solid #2B2F3B; border-top: 3px solid #FFC83D; border-radius: 14px;
               padding: 18px; min-height: 108px; transition: transform .15s, border-color .15s; }
.metric-card:hover { transform: translateY(-3px); border-color: #FFC83D; }
.stApp .metric-label { color: #E6E8F0; font-size: 13px; font-weight: 500; }
.stApp .metric-value { color: #FFFFFF; font-size: 27px; font-weight: 700; margin-top: 8px; }
.stApp .metric-note { color: #CDD0DC; font-size: 12px; margin-top: 4px; }
.info-card { background: #12141C; border: 1px solid #2B2F3B; border-radius: 14px; padding: 15px; margin-bottom: 12px; }
.stApp .info-label { color: #CDD0DC; font-size: 12px; }
.stApp .info-value { font-size: 16px; font-weight: 600; color: #FFFFFF; margin-top: 6px; word-break: break-word; }
.section-title { font-family: 'Fraunces', serif; font-size: 28px; margin: 34px 0 6px; color: #FFFFFF; }
.stApp .section-sub { color: #DADDE8; font-size: 14px; margin-bottom: 14px; }
.stApp .spot-tag { color: #FFC83D; font-size: 12px; font-weight: 700; letter-spacing: .16em; }
.spot-title { font-family: 'Fraunces', serif; font-size: 36px; font-weight: 800; margin: 8px 0; line-height: 1.1; color: #FFFFFF; }
.feature { background: #12141C; border: 1px solid #2B2F3B; border-radius: 16px; padding: 20px; min-height: 110px; }
.feature b { font-size: 17px; color: #FFFFFF; }
.stApp .feature div { color: #DADDE8; font-size: 14px; margin-top: 6px; }

.poster { border-radius: 12px; display: flex; align-items: center; justify-content: center; color: rgba(255,255,255,.95);
          font-family: 'Fraunces', serif; font-weight: 800; font-size: 44px; border: 1px solid #2F3342;
          background-size: cover; background-position: center; }
.stApp .mtitle { font-weight: 700; font-size: 16px; margin: 10px 0 2px; line-height: 1.25; min-height: 40px; color: #FFFFFF; }
.stApp .mmeta { color: #E6E8F0; font-size: 13px; }
.stApp .mgenre { color: #D8D6CA; font-size: 12px; margin: 4px 0 8px; min-height: 18px; }
.stApp .star { color: #FFC83D; font-weight: 700; }

.stButton button, .stDownloadButton button { border-radius: 10px; border: 1px solid #FFC83D; background: transparent; font-weight: 700; }
.stButton button p, .stDownloadButton button p { color: #FFC83D !important; }
.stButton button:hover, .stDownloadButton button:hover { background: #FFC83D; border-color: #FFC83D; }
.stButton button:hover p, .stDownloadButton button:hover p { color: #15120A !important; }
[data-testid="stDataFrame"] { border: 1px solid #2B2F3B; border-radius: 12px; }
.stApp .footer { text-align: center; color: #CDD0DC; font-size: 12px; padding: 30px 0 6px; }
@media (max-width: 700px) { .hero { padding: 40px 22px; } .hero-title { font-size: 36px; } }
</style>
""", unsafe_allow_html=True)

# ------------------------------------------------------------------
# DATA LOADING
# ------------------------------------------------------------------
# Always resolve paths relative to this app.py file, not the terminal's current folder.
APP_DIR = os.path.dirname(os.path.abspath(__file__))

SEARCH_DIRS = []
for _d in (APP_DIR, os.path.join(APP_DIR, "data"), os.getcwd()):
    if os.path.isdir(_d) and _d not in SEARCH_DIRS:
        SEARCH_DIRS.append(_d)

ALIASES = {
    "movie_id": ["movie_id", "movieid", "id"],
    "title": ["title", "movie_title", "original_title", "name"],
    "release_date": ["release_date", "releasedate", "release", "date"],
    "status": ["status"], "budget": ["budget"],
    "revenue": ["revenue", "gross", "box_office"],
    "runtime": ["runtime", "duration", "run_time"],
    "genres": ["genres", "genre"],
    "production": ["production_companies", "production", "company"],
    "countries": ["production_countries", "countries", "country"],
    "spoken_language": ["spoken_languages", "spoken_language", "language"],
    "original_language": ["original_language", "original_lang"],
    "vote_average": ["vote_average", "vote_avg", "rating", "average_rating"],
    "vote_count": ["vote_count", "votes", "number_of_votes"],
    "popularity": ["popularity"],
    "adult": ["adult", "is_adult"],
    "overview": ["overview", "description", "plot", "summary"],
    "poster": ["poster", "poster_url", "poster_path", "image"],
}
MULTI = ["genres", "production", "countries", "spoken_language"]   # values separated by | in your Excel

LANG_NAMES = {
    "en": "English", "es": "Spanish", "fr": "French", "de": "German", "it": "Italian", "pt": "Portuguese",
    "ru": "Russian", "ja": "Japanese", "ko": "Korean", "zh": "Chinese", "cn": "Cantonese", "hi": "Hindi",
    "ta": "Tamil", "te": "Telugu", "ml": "Malayalam", "kn": "Kannada", "bn": "Bengali", "mr": "Marathi",
    "pa": "Punjabi", "ur": "Urdu", "tr": "Turkish", "ar": "Arabic", "fa": "Persian", "he": "Hebrew",
    "th": "Thai", "vi": "Vietnamese", "id": "Indonesian", "tl": "Tagalog", "ms": "Malay", "sv": "Swedish",
    "da": "Danish", "no": "Norwegian", "nb": "Norwegian", "fi": "Finnish", "nl": "Dutch", "pl": "Polish",
    "cs": "Czech", "hu": "Hungarian", "ro": "Romanian", "el": "Greek", "uk": "Ukrainian", "bg": "Bulgarian",
    "sr": "Serbian", "hr": "Croatian", "sk": "Slovak", "sl": "Slovenian", "lt": "Lithuanian", "lv": "Latvian",
    "et": "Estonian", "is": "Icelandic", "ca": "Catalan", "eu": "Basque", "gl": "Galician", "xx": "No language",
}
MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September",
          "October", "November", "December"]


def _usable(files):
    return [f for f in files if not os.path.basename(f).startswith("~$")
            and "cinescope_filtered" not in os.path.basename(f).lower()]


def find_dataset():
    """Find the movie workbook beside app.py, in its data folder, or in Downloads."""
    if DATA_PATH.strip():
        candidate = os.path.expandvars(os.path.expanduser(DATA_PATH.strip().strip('"')))
        if os.path.isfile(candidate):
            return os.path.abspath(candidate)
        raise FileNotFoundError(f"DATA_PATH does not exist: {candidate}")

    found = []
    # First check the project folder and its optional data subfolder.
    for folder in SEARCH_DIRS:
        if os.path.isdir(folder):
            for pattern in ("*.xlsx", "*.xls", "*.csv"):
                found.extend(glob.glob(os.path.join(folder, pattern)))

    found = _usable(found)

    # If the workbook is in Downloads rather than beside app.py, search a few
    # common user folders (limited depth so startup remains quick).
    if not found:
        home = os.path.expanduser("~")
        bases = [
            os.path.join(home, "Downloads"),
            os.path.join(home, "Desktop"),
            os.path.join(home, "Documents"),
            os.path.join(home, "OneDrive", "Desktop"),
            os.path.join(home, "OneDrive", "Documents"),
        ]
        for base in bases:
            if not os.path.isdir(base):
                continue
            for root, dirs, files in os.walk(base):
                depth = os.path.relpath(root, base).count(os.sep)
                if depth >= 3:
                    dirs[:] = []
                for name in files:
                    lower = name.lower()
                    if lower.endswith((".xlsx", ".xls", ".csv")) and any(
                        key in lower for key in ("movie", "khushi", "cinescope")
                    ):
                        found.append(os.path.join(root, name))
        found = _usable(found)

    if not found:
        raise FileNotFoundError(
            "No movie Excel/CSV file found. Put 'khushi data (1).xlsx' in the same folder as app.py."
        )

    def rank(path):
        name = os.path.basename(path).lower()
        # Prefer the user's named workbook, then the newest movie-like file.
        priority = 0 if "khushi data" in name else 1 if "khushi" in name else 2 if "movie" in name else 3
        return (priority, -os.path.getmtime(path))

    found.sort(key=rank)
    return os.path.abspath(found[0])


def clean_number(series):
    s = series.astype(str).str.replace(r"[,$₹€]", "", regex=True).str.strip()
    s = s.replace({"": np.nan, "nan": np.nan, "None": np.nan, "null": np.nan, "N/A": np.nan})
    return pd.to_numeric(s, errors="coerce")


def split_multi(series):
    """'Drama|Thriller' -> ['Drama', 'Thriller'] (falls back to commas if there is no | in the column)."""
    s = series.fillna("").astype(str)
    sep = "|" if s.str.contains("|", regex=False).any() else ","

    def one(v):
        out = []
        for p in v.split(sep):
            p = p.strip()
            if p and p.lower() not in ("unknown", "nan", "none", "null") and p not in out:
                out.append(p)
        return out
    return s.map(one)


def prepare_data(raw):
    df = raw.copy()

    # Normalize both incoming column names and aliases in exactly the same way.
    def normalize_column(value):
        return re.sub(r"[^a-z0-9]", "", str(value).strip().lower())

    df.columns = [normalize_column(c) for c in df.columns]
    rename = {}
    for std, choices in ALIASES.items():
        normalized_choices = [normalize_column(c) for c in choices]
        for candidate in normalized_choices:
            if candidate in df.columns and candidate not in rename:
                rename[candidate] = std
                break
    df = df.rename(columns=rename)
    df = df.loc[:, ~df.columns.duplicated()].reset_index(drop=True)
    df["_rid"] = np.arange(len(df))

    if "title" not in df.columns:
        df["title"] = "Movie " + df["_rid"].astype(str)

    for c in MULTI:                              # list columns for analysis + readable text for display
        if c in df.columns:
            lst = split_multi(df[c])
            df[c + "_list"] = lst
            df[c] = lst.map(lambda x: ", ".join(x) if x else "Unknown")

    for c in ["title", "status", "original_language", "overview"]:
        if c in df.columns:
            df[c] = df[c].fillna("Unknown").astype(str).str.strip()
            df.loc[df[c].str.lower().isin(["", "nan", "none", "null"]), c] = "Unknown"
    if "original_language" in df.columns:
        df["language_name"] = df["original_language"].map(
            lambda c: LANG_NAMES.get(str(c).lower(), str(c).upper() if c != "Unknown" else "Unknown"))

    for c in ["budget", "revenue", "runtime", "vote_average", "vote_count", "popularity"]:
        if c in df.columns:
            df[c] = clean_number(df[c])

    if "release_date" in df.columns:
        original_dates = df["release_date"]

        # The uploaded workbook stores real Excel dates. For text dates such as
        # 01-01-2010, dayfirst=True prevents month/day confusion. Numeric Excel
        # serial dates are handled separately to avoid Unix-epoch (1970) dates.
        if pd.api.types.is_datetime64_any_dtype(original_dates):
            parsed = pd.to_datetime(original_dates, errors="coerce")
        elif pd.api.types.is_numeric_dtype(original_dates):
            numeric_dates = pd.to_numeric(original_dates, errors="coerce")
            parsed = pd.to_datetime(numeric_dates, unit="D", origin="1899-12-30", errors="coerce")
            parsed = parsed.where(numeric_dates.between(20000, 80000))
        else:
            parsed = pd.to_datetime(original_dates, dayfirst=True, errors="coerce")

        df["release_date"] = parsed
        df["release_year"] = parsed.dt.year.astype("Int64")
        df["release_month"] = parsed.dt.month_name()
    else:
        df["release_year"] = np.nan
        df["release_month"] = np.nan

    if "adult" in df.columns:
        df["adult_label"] = (df["adult"].astype(str).str.lower().str.strip()
                             .map({"true": "Adult", "1": "Adult", "yes": "Adult",
                                   "false": "Non-Adult", "0": "Non-Adult", "no": "Non-Adult"}).fillna("Unknown"))
    if {"budget", "revenue"} <= set(df.columns):
        df["profit"] = df["revenue"] - df["budget"]
    return df


@st.cache_data(show_spinner="Loading CineScope movies...")
def load_prepared(path, modified):
    if path.lower().endswith((".xlsx", ".xls")):
        raw = pd.read_excel(path, engine="openpyxl")
    else:
        try:
            raw = pd.read_csv(path, encoding="utf-8-sig")
        except UnicodeDecodeError:
            raw = pd.read_csv(path, encoding="latin-1")
    return prepare_data(raw)


try:
    DATA_FILE = find_dataset()
    df = load_prepared(DATA_FILE, os.path.getmtime(DATA_FILE))
    source_label = os.path.basename(DATA_FILE)
except Exception as exc:
    st.error(f"Could not load the movie data: {exc}")
    st.markdown("*App is running from this folder:*")
    st.code(APP_DIR)
    try:
        st.markdown("*Files in this folder:*")
        st.code("\n".join(os.listdir(APP_DIR)) or "(empty)")
    except Exception:
        pass
    st.info("Put your Excel file in the folder above, then refresh the page.")
    st.stop()

if df.empty:
    st.error("The dataset has no rows.")
    st.stop()

if "release_date" in df.columns and df["release_year"].notna().sum() == 0:
    st.warning("Release dates could not be read. Check the release_date column in your Excel file.")

# ------------------------------------------------------------------
# HELPERS
# ------------------------------------------------------------------
has = lambda c: c in df.columns


def stretch(fn, *args, **kwargs):
    """Full-width widgets on both old and new Streamlit versions."""
    try:
        return fn(*args, width="stretch", **kwargs)
    except Exception:
        return fn(*args, use_container_width=True, **kwargs)


def money(v):
    if pd.isna(v):
        return "N/A"
    v = float(v)
    sign = "-" if v < 0 else ""
    v = abs(v)
    for limit, suffix in ((1e9, "B"), (1e6, "M"), (1e3, "K")):
        if v >= limit:
            return f"{sign}${v / limit:.2f}{suffix}"
    return f"{sign}${v:,.0f}"


def explode(frame, col, name):
    if col not in frame.columns:
        return pd.DataFrame(columns=[name])
    out = frame.explode(col).rename(columns={col: name})
    return out[out[name].notna() & (out[name] != "")]


def top_with_other(counts, n=8):
    """Keep the top n slices of a count Series and group the rest as 'Other' (for readable pie charts)."""
    counts = counts.sort_values(ascending=False)
    if len(counts) > n:
        counts = pd.concat([counts.head(n), pd.Series({"Other": counts.iloc[n:].sum()})])
    return counts


def rated(frame):
    return frame[frame["vote_average"] > 0] if "vote_average" in frame.columns else frame.head(0)


GENRES = explode(df, "genres_list", "g")["g"].value_counts() if has("genres_list") else pd.Series(dtype=int)


def metric_card(label, value, note=""):
    st.markdown(f'<div class="metric-card"><div class="metric-label">{label}</div>'
                f'<div class="metric-value">{value}</div><div class="metric-note">{note}</div></div>',
                unsafe_allow_html=True)


def section(title, sub=""):
    st.markdown(f'<div class="section-title">{title}</div>', unsafe_allow_html=True)
    if sub:
        st.markdown(f'<div class="section-sub">{sub}</div>', unsafe_allow_html=True)


def show(fig, height=380):
    fig.update_layout(template="plotly_dark", height=height, paper_bgcolor="rgba(0,0,0,0)",
                      plot_bgcolor="rgba(0,0,0,0)", font=dict(family="Outfit", color=WHITE, size=13),
                      margin=dict(l=10, r=10, t=55, b=20), colorway=SEQ,
                      title=dict(font=dict(size=18, color=WHITE)),
                      legend=dict(font=dict(color=WHITE)), hoverlabel=dict(font=dict(color="#111")))
    fig.update_xaxes(gridcolor="#262A36", tickfont=dict(color=WHITE), title_font=dict(color=WHITE))
    fig.update_yaxes(gridcolor="#262A36", tickfont=dict(color=WHITE), title_font=dict(color=WHITE))
    try:
        st.plotly_chart(fig, width="stretch")
    except Exception:
        st.plotly_chart(fig, use_container_width=True)


def pie(counts, title, hole=0.45, n=8):
    c = top_with_other(counts, n).reset_index()
    c.columns = ["Label", "Count"]
    fig = px.pie(c, names="Label", values="Count", hole=hole, title=title)
    fig.update_traces(textposition="inside", textinfo="percent+label", textfont=dict(color="#101010", size=13),
                      marker=dict(line=dict(color="#0A0B10", width=2)))
    fig.update_layout(showlegend=False)
    return fig


def poster_html(row, height=200):
    title = str(row.get("title", "?"))
    url = str(row.get("poster", "") or "")
    if url.startswith("/"):
        url = "https://image.tmdb.org/t/p/w500" + url
    if url.startswith("http"):
        return (f'<div class="poster" style="height:{height}px;'
                f'background-image:url(\'{html.escape(url, quote=True)}\')"></div>')
    hue = (sum(ord(ch) for ch in title) * 37) % 360
    initials = "".join(w[0] for w in re.findall(r"[A-Za-z0-9]+", title)[:2]).upper() or "🎬"
    return (f'<div class="poster" style="height:{height}px;background:linear-gradient(150deg,'
            f'hsl({hue},55%,32%),hsl({(hue + 40) % 360},60%,14%))">{html.escape(initials)}</div>')


def year_text(row):
    y = row.get("release_year")
    return str(int(y)) if pd.notna(y) else "Year n/a"


def rating_text(row):
    r = row.get("vote_average")
    return f'<span class="star">★ {r:.1f}</span>' if pd.notna(r) and r > 0 else "Not rated"


def featured(frame, by, n=8, min_votes=True):
    frame = frame.dropna(subset=[by])
    if by == "vote_average":
        frame = frame[frame["vote_average"] > 0]
    if min_votes and "vote_count" in frame.columns and frame["vote_count"].notna().any():
        frame = frame[frame["vote_count"] >= frame["vote_count"].median()]
    return frame.sort_values(by, ascending=False).head(n)


def open_movie(rid, key, label="Open movie"):
    if stretch(st.button, label, key=key):
        st.session_state["movie"] = int(rid)
        st.query_params["movie"] = str(int(rid))
        st.switch_page(PAGE_MOVIE)


def movie_card(row, key):
    glist = row.get("genres_list", []) if "genres_list" in row.index else []
    genres = html.escape(", ".join(glist[:2]))
    st.markdown(f'{poster_html(row)}<div class="mtitle">{html.escape(str(row["title"]))}</div>'
                f'<div class="mmeta">{year_text(row)} &nbsp;·&nbsp; {rating_text(row)}</div>'
                f'<div class="mgenre">{genres}</div>', unsafe_allow_html=True)
    open_movie(row["_rid"], key)


def movie_grid(frame, prefix, per_row=4):
    for r_i in range(0, len(frame), per_row):
        cols = st.columns(per_row)
        for c_i, (_, row) in enumerate(frame.iloc[r_i:r_i + per_row].iterrows()):
            with cols[c_i]:
                movie_card(row, f"{prefix}{r_i}{c_i}_{row['_rid']}")
        st.write("")


# ------------------------------------------------------------------
# SIDEBAR FILTERS (apply to every page)
# ------------------------------------------------------------------
filtered = df.copy()
with st.sidebar:
    st.markdown("## 🎬 CineScope")
    st.caption("Filters apply to every page")
    st.divider()
    if df["release_year"].notna().any():
        years = sorted(df["release_year"].dropna().astype(int).unique())
        if len(years) > 1:
            yr = st.slider("Release year", years[0], years[-1], (years[0], years[-1]))
            filtered = filtered[filtered["release_year"].between(*yr)]
    if len(GENRES):
        sel = st.multiselect("Genre", sorted(GENRES.index))
        if sel:
            chosen = set(sel)
            filtered = filtered[filtered["genres_list"].map(lambda l: bool(chosen & set(l)))]
    if has("status"):
        sel = st.multiselect("Status", sorted(df["status"].unique()))
        if sel:
            filtered = filtered[filtered["status"].isin(sel)]
    if has("language_name"):
        sel = st.multiselect("Original language", sorted(df["language_name"].unique()))
        if sel:
            filtered = filtered[filtered["language_name"].isin(sel)]
    if has("countries_list"):
        all_c = sorted(explode(df, "countries_list", "c")["c"].unique())
        sel = st.multiselect("Production country", all_c)
        if sel:
            chosen = set(sel)
            filtered = filtered[filtered["countries_list"].map(lambda l: bool(chosen & set(l)))]
    st.divider()
    st.caption(f"Showing {len(filtered):,} of {len(df):,} movies")


def ready():
    if filtered.empty:
        st.warning("No movies match these filters. Open the sidebar (top-left arrow) and widen them.")
        return False
    return True


# ------------------------------------------------------------------
# HOME
# ------------------------------------------------------------------
def page_home():
    if not ready():
        return
    st.markdown("""
    <div class="hero">
        <div class="eyebrow">MOVIE DISCOVERY &amp; ANALYTICS</div>
        <div class="hero-title">Find your next movie. <span>See the numbers</span> behind every hit.</div>
        <div class="hero-sub">Browse every title from 2010 to 2027, open any movie on its own page, and explore
        ratings, genres, budgets and box-office trends in one place.</div>
    </div>""", unsafe_allow_html=True)

    with st.container(key="cta"):
        a, b, c, _ = st.columns([1, 1, 1, 2])
        stretch(a.page_link, PAGE_BROWSE, label="Browse movies", icon="🎞️")
        stretch(b.page_link, PAGE_ANALYTICS, label="Analytics", icon="📊")
        stretch(c.page_link, PAGE_PERFORMANCE, label="Performance", icon="📈")
    st.write("")

    rt = rated(filtered)
    avg = rt["vote_average"].mean() if has("vote_average") and not rt.empty else np.nan
    yrs = filtered["release_year"].dropna()
    span = f"{int(yrs.min())}–{int(yrs.max())}" if not yrs.empty else "N/A"
    n_countries = explode(filtered, "countries_list", "c")["c"].nunique() if has("countries_list") else 0
    k = st.columns(5)
    with k[0]: metric_card("Movies", f"{len(filtered):,}", "In current selection")
    with k[1]: metric_card("Years covered", span, "Release range")
    with k[2]: metric_card("Average rating", f"{avg:.2f}" if pd.notna(avg) else "N/A", "Out of 10 (rated movies)")
    with k[3]: metric_card("Total revenue", money(filtered["revenue"].sum(min_count=1)) if has("revenue") else "N/A", "Dataset currency")
    with k[4]: metric_card("Countries", f"{n_countries:,}", "Production countries")
    st.write("")
    k = st.columns(5)
    with k[0]: metric_card("Total budget", money(filtered["budget"].sum(min_count=1)) if has("budget") else "N/A", "All movies")
    rtm = filtered[filtered["runtime"] > 0]["runtime"].mean() if has("runtime") else np.nan
    with k[1]: metric_card("Average runtime", f"{rtm:.0f} min" if pd.notna(rtm) else "N/A", "Movies with runtime")
    with k[2]: metric_card("Total votes", f"{filtered['vote_count'].sum():,.0f}" if has("vote_count") else "N/A", "All movies")
    with k[3]: metric_card("Languages", f"{filtered['language_name'].nunique():,}" if has("language_name") else "N/A", "Original languages")
    with k[4]: metric_card("Genres", f"{explode(filtered, 'genres_list', 'g')['g'].nunique():,}" if has("genres_list") else "N/A", "Unique genres")

    if has("vote_average"):
        top = featured(filtered, "vote_average", 1)
        if not top.empty:
            row = top.iloc[0]
            section("Spotlight")
            with st.container(key="spot"):
                l, r = st.columns([1, 2.6])
                with l:
                    st.markdown(poster_html(row, 260), unsafe_allow_html=True)
                with r:
                    text = ""
                    if "overview" in row.index and row["overview"] != "Unknown":
                        text = f'<div class="spot-text">{html.escape(row["overview"][:300])}</div>'
                    votes = (f' &nbsp;·&nbsp; {row["vote_count"]:,.0f} votes'
                             if has("vote_count") and pd.notna(row["vote_count"]) else "")
                    st.markdown(f'<div class="spot-tag">TOP RATED</div>'
                                f'<div class="spot-title">{html.escape(str(row["title"]))}</div>'
                                f'<div class="mmeta">{year_text(row)} &nbsp;·&nbsp; {rating_text(row)}{votes}</div>'
                                f'<div class="mgenre">{html.escape(str(row.get("genres", "")))}</div>{text}',
                                unsafe_allow_html=True)
                    st.write("")
                    open_movie(row["_rid"], "spotlight", "Open this movie")

    if len(GENRES):
        section("Explore by genre", "Pick a genre to see its movies.")
        top_genres = list(GENRES.head(8).items())
        for i in range(0, len(top_genres), 4):
            cols = st.columns(4)
            for j, (g, n) in enumerate(top_genres[i:i + 4]):
                if stretch(cols[j].button, f"{g}  ·  {n:,}", key=f"genre_{g}"):
                    st.session_state["pending_genre"] = g
                    st.switch_page(PAGE_BROWSE)

    if has("vote_average"):
        section("Top rated", "Movies with very few votes are left out.")
        movie_grid(featured(filtered, "vote_average", 8), "home_top")
    pop = "popularity" if has("popularity") else ("vote_count" if has("vote_count") else None)
    if pop:
        section("Trending now", "Most popular titles in the selection.")
        movie_grid(featured(filtered, pop, 4, min_votes=False), "home_pop")
    if has("release_date"):
        released = filtered[filtered["status"] == "Released"] if has("status") else filtered
        section("Latest releases")
        movie_grid(released.dropna(subset=["release_date"]).sort_values("release_date", ascending=False).head(4), "home_new")
        if has("status"):
            up = filtered[filtered["status"] != "Released"].dropna(subset=["release_date"])
            if not up.empty:
                section("Coming soon", "Planned and in post-production titles.")
                movie_grid(up.sort_values("release_date").head(4), "home_up")

    section("The big picture")
    l, r = st.columns(2)
    with l:
        if not yrs.empty:
            yd = filtered.dropna(subset=["release_year"]).groupby("release_year").size().reset_index(name="Movies")
            fig = px.area(yd, x="release_year", y="Movies", title="Movies released per year")
            fig.update_traces(line_color=GOLD, fillcolor="rgba(255,200,61,.20)")
            show(fig)
    with r:
        gdf = explode(filtered, "genres_list", "genre")
        if not gdf.empty:
            gc = gdf["genre"].value_counts().head(10).reset_index()
            gc.columns = ["Genre", "Movies"]
            fig = px.bar(gc, x="Movies", y="Genre", orientation="h", title="Top 10 genres", color_discrete_sequence=[TEAL])
            fig.update_layout(yaxis={"categoryorder": "total ascending"})
            show(fig)

    section("Go deeper")
    f1, f2, f3 = st.columns(3)
    for col, icon, name, page, text in (
            (f1, "📊", "Analytics", PAGE_ANALYTICS, "Genres, ratings, runtimes, languages, countries and studios."),
            (f2, "📈", "Performance", PAGE_PERFORMANCE, "Yearly trends, budget vs revenue, hits and flops."),
            (f3, "🗂️", "Data", PAGE_DATA, "Browse the full table and download your selection.")):
        with col:
            st.markdown(f'<div class="feature"><b>{icon} {name}</b><div>{text}</div></div>', unsafe_allow_html=True)
            st.page_link(page, label=f"Open {name}", icon="➡️")


# ------------------------------------------------------------------
# BROWSE
# ------------------------------------------------------------------
def page_browse():
    if not ready():
        return
    st.markdown("## Browse movies")
    if "pending_genre" in st.session_state:
        st.session_state["browse_genre"] = st.session_state.pop("pending_genre")

    a, b, c = st.columns([3, 1.4, 1.4])
    q = a.text_input("Search by title", value=st.query_params.get("q", ""), placeholder="Type a movie name")
    genre = b.selectbox("Genre", ["All"] + sorted(GENRES.index), key="browse_genre") if len(GENRES) else "All"
    sorts = {"Title A to Z": ("title", True)}
    if has("vote_average"): sorts = {"Highest rated": ("vote_average", False), **sorts}
    if has("vote_count"): sorts["Most voted"] = ("vote_count", False)
    if has("popularity"): sorts["Most popular"] = ("popularity", False)
    if has("release_date"):
        sorts["Newest"] = ("release_date", False)
        sorts["Oldest"] = ("release_date", True)
    if has("revenue"): sorts["Highest revenue"] = ("revenue", False)
    choice = c.selectbox("Sort by", list(sorts))

    if q:
        st.query_params["q"] = q
    elif "q" in st.query_params:
        del st.query_params["q"]

    res = filtered
    if q:
        res = res[res["title"].str.contains(q, case=False, na=False, regex=False)]
    if genre != "All":
        res = res[res["genres_list"].map(lambda l: genre in l)]
    col, asc = sorts[choice]
    res = res.sort_values(col, ascending=asc, na_position="last")

    if res.empty:
        st.info("No movies match that search. Try a shorter name or another genre.")
        return
    pages = max(1, -(-len(res) // PAGE_SIZE))
    page = int(st.number_input("Page", 1, pages, 1)) if pages > 1 else 1
    start = (page - 1) * PAGE_SIZE
    st.caption(f"Showing {start + 1}-{min(start + PAGE_SIZE, len(res))} of {len(res):,} movies")
    movie_grid(res.iloc[start:start + PAGE_SIZE], f"browse{page}")


# ------------------------------------------------------------------
# MOVIE (own page)
# ------------------------------------------------------------------
def page_movie():
    rid = st.session_state.get("movie")
    qp = st.query_params.get("movie", "")
    if rid is None and str(qp).isdigit():
        rid = int(qp)
    row = None
    if rid is not None:
        match = df[df["_rid"] == rid]
        row = match.iloc[0] if not match.empty else None

    if row is None:
        st.markdown("## Find a movie")
        q = st.text_input("Search by title", placeholder="Type a movie name")
        res = df[df["title"].str.contains(q, case=False, na=False, regex=False)] if q else df.head(0)
        if q and res.empty:
            st.info("No titles match that search.")
        for _, r in res.head(15).iterrows():
            open_movie(r["_rid"], f"find{r['_rid']}", f"{r['title']} ({year_text(r)})")
        if not q:
            st.caption("Or open a movie from the Browse page.")
        return

    st.session_state["movie"] = int(row["_rid"])
    st.query_params["movie"] = str(int(row["_rid"]))
    st.page_link(PAGE_BROWSE, label="Back to browse", icon="⬅️")

    left, right = st.columns([1, 2.4])
    with left:
        st.markdown(poster_html(row, 360), unsafe_allow_html=True)
    with right:
        st.markdown(f'<div class="hero-title" style="font-size:40px;margin-top:0">{html.escape(str(row["title"]))}</div>',
                    unsafe_allow_html=True)
        chips = [year_text(row)]
        if has("runtime") and pd.notna(row["runtime"]) and row["runtime"] > 0:
            chips.append(f'{row["runtime"]:.0f} min')
        if has("status") and row["status"] != "Unknown":
            chips.append(str(row["status"]))
        chips += list(row.get("genres_list", []))[:4]
        st.markdown('<div class="chips">' + "".join(f'<span class="chip">{html.escape(c)}</span>' for c in chips)
                    + "</div>", unsafe_allow_html=True)
        if has("overview") and row["overview"] != "Unknown":
            st.write("")
            st.write(row["overview"])
        st.write("")
        k = st.columns(3)
        has_rating = has("vote_average") and pd.notna(row["vote_average"]) and row["vote_average"] > 0
        with k[0]: metric_card("Rating", f'{row["vote_average"]:.1f} / 10' if has_rating else "Not rated",
                               f'{row["vote_count"]:,.0f} votes' if has("vote_count") and pd.notna(row["vote_count"]) else "")
        with k[1]: metric_card("Budget", money(row["budget"]) if has("budget") else "N/A", "Dataset currency")
        with k[2]: metric_card("Revenue", money(row["revenue"]) if has("revenue") else "N/A",
                               f'Profit {money(row["profit"])}' if has("profit") and pd.notna(row.get("profit")) and row["revenue"] > 0 else "")

    section("Details")
    fields = [("Movie ID", "movie_id"), ("Release date", "release_date"), ("Status", "status"),
              ("Production companies", "production"), ("Countries", "countries"),
              ("Spoken languages", "spoken_language"), ("Original language", "language_name"),
              ("Popularity", "popularity"), ("Adult", "adult_label")]
    dc = st.columns(3)
    for i, (lbl, col) in enumerate([(l, c) for l, c in fields if has(c)]):
        v = row[col]
        if pd.isna(v):
            v = "Not available"
        elif col == "release_date":
            v = v.strftime("%d %b %Y")
        elif col == "popularity":
            v = f"{v:.1f}"
        elif col == "movie_id":
            v = f"{int(v)}"
        with dc[i % 3]:
            st.markdown(f'<div class="info-card"><div class="info-label">{lbl}</div>'
                        f'<div class="info-value">{html.escape(str(v))}</div></div>', unsafe_allow_html=True)

    if has_rating:
        fig = px.histogram(rated(df), x="vote_average", nbins=25, title="How this rating compares",
                           labels={"vote_average": "Rating"})
        fig.update_traces(marker_color="#5A5E70")
        fig.add_vline(x=float(row["vote_average"]), line_color=GOLD, line_width=3,
                      annotation_text="This movie", annotation_font_color=GOLD)
        show(fig, 320)

    glist = list(row.get("genres_list", []))
    if glist:
        first = glist[0]
        similar = df[(df["_rid"] != row["_rid"]) & df["genres_list"].map(lambda l: first in l)]
        similar = featured(similar, "vote_average", 4)
        if not similar.empty:
            section(f"More {first} movies")
            movie_grid(similar, "similar")


# ------------------------------------------------------------------
# ANALYTICS
# ------------------------------------------------------------------
def page_analytics():
    if not ready():
        return
    st.markdown("## Analytics")
    gdf = explode(filtered, "genres_list", "genre")
    t_gen, t_rate, t_lang, t_studio, t_corr = st.tabs(
        ["🎭 Genres", "⭐ Ratings & runtime", "🌍 Languages & countries", "🏢 Studios", "🔗 Correlations"])

    # ---- Genres
    with t_gen:
        if gdf.empty:
            st.info("No genre data in this selection.")
        else:
            l, r = st.columns(2)
            with l:
                show(pie(gdf["genre"].value_counts(), "Genre share", n=12), 420)
            with r:
                gc = gdf["genre"].value_counts().reset_index()
                gc.columns = ["Genre", "Movies"]
                fig = px.bar(gc, x="Movies", y="Genre", orientation="h", title="Movies per genre",
                             color_discrete_sequence=[GOLD])
                fig.update_layout(yaxis={"categoryorder": "total ascending"})
                show(fig, 420)
            l, r = st.columns(2)
            with l:
                gr = rated(gdf).groupby("genre")["vote_average"].mean().sort_values(ascending=False).reset_index()
                gr.columns = ["Genre", "Average rating"]
                fig = px.bar(gr, x="Genre", y="Average rating", title="Average rating by genre",
                             color_discrete_sequence=[TEAL])
                fig.update_yaxes(range=[max(0, gr["Average rating"].min() - 0.5), gr["Average rating"].max() + 0.2])
                show(fig)
            with r:
                if has("revenue"):
                    rv = gdf.groupby("genre")["revenue"].sum().sort_values(ascending=False).reset_index()
                    rv.columns = ["Genre", "Total revenue"]
                    show(px.bar(rv, x="Genre", y="Total revenue", title="Total revenue by genre",
                                color_discrete_sequence=[ROSE]))
            top_g = gdf["genre"].value_counts().head(10).index
            hm = (gdf[gdf["genre"].isin(top_g)].dropna(subset=["release_year"])
                  .assign(release_year=lambda d: d["release_year"].astype(int))
                  .groupby(["genre", "release_year"]).size().unstack(fill_value=0))
            if not hm.empty:
                fig = px.imshow(hm, aspect="auto", color_continuous_scale="YlOrRd",
                                title="Top genres by release year (movie count)",
                                labels=dict(x="Year", y="Genre", color="Movies"))
                show(fig, 420)
            l, r = st.columns(2)
            with l:
                gy = (gdf[gdf["genre"].isin(top_g[:6])].dropna(subset=["release_year"])
                      .groupby(["release_year", "genre"]).size().reset_index(name="Movies"))
                show(px.line(gy, x="release_year", y="Movies", color="genre", markers=True,
                             title="Top 6 genres over time"))
            with r:
                show(px.treemap(gdf["genre"].value_counts().reset_index().rename(columns={"genre": "Genre", "count": "Movies"})
                                if "count" in gdf["genre"].value_counts().reset_index().columns
                                else gdf["genre"].value_counts().reset_index(name="Movies").rename(columns={"index": "Genre"}),
                                path=["Genre"], values="Movies", title="Genre treemap"))

    # ---- Ratings & runtime
    with t_rate:
        rt = rated(filtered)
        l, r = st.columns(2)
        with l:
            if not rt.empty:
                fig = px.histogram(rt, x="vote_average", nbins=30, title="Rating distribution", labels={"vote_average": "Rating"})
                fig.update_traces(marker_color=GOLD)
                show(fig)
        with r:
            if not rt.empty:
                bins = pd.cut(rt["vote_average"], [0, 4, 5, 6, 7, 8, 10.01],
                              labels=["Below 4", "4 – 5", "5 – 6", "6 – 7", "7 – 8", "8 and above"], right=False)
                show(pie(bins.value_counts(), "Rating groups", n=10))
        rtm = filtered[filtered["runtime"] > 0] if has("runtime") else filtered.head(0)
        l, r = st.columns(2)
        with l:
            if not rtm.empty:
                fig = px.histogram(rtm[rtm["runtime"] <= 240], x="runtime", nbins=40, title="Runtime distribution (up to 240 min)",
                                   labels={"runtime": "Minutes"})
                fig.update_traces(marker_color=SKY)
                show(fig)
        with r:
            if not rtm.empty:
                rb = pd.cut(rtm["runtime"], [0, 60, 90, 120, 150, 180, 1000],
                            labels=["Under 60", "60 – 90", "90 – 120", "120 – 150", "150 – 180", "180+"], right=False)
                show(pie(rb.value_counts(), "Runtime groups (minutes)", n=10))
        vr = rt.dropna(subset=["vote_count"]) if has("vote_count") else rt.head(0)
        if not vr.empty:
            if len(vr) > 5000:
                vr = vr.sample(5000, random_state=1)
            fig = px.scatter(vr, x="vote_count", y="vote_average", opacity=0.6, hover_name="title", log_x=True,
                             color="vote_average", color_continuous_scale="YlOrRd", title="Votes vs rating (log scale)")
            fig.update_layout(coloraxis_showscale=False)
            show(fig, 420)
        l, r = st.columns(2)
        with l:
            if has("status") and not rt.empty:
                show(px.box(rt, x="status", y="vote_average", color="status", title="Rating by status"))
        with r:
            if has("language_name") and not rt.empty:
                top_l = filtered["language_name"].value_counts().head(8).index
                show(px.box(rt[rt["language_name"].isin(top_l)], x="language_name", y="vote_average",
                            color="language_name", title="Rating by top languages"))
        if has("status"):
            l, r = st.columns(2)
            with l:
                show(pie(filtered["status"].value_counts(), "Release status", n=10))
            with r:
                if has("adult_label") and filtered["adult_label"].nunique() > 1:
                    show(pie(filtered["adult_label"].value_counts(), "Adult classification"))
                else:
                    st.markdown("&nbsp;")
                    st.info("Every movie in the current selection has the same adult classification "
                            f"({filtered['adult_label'].iloc[0] if has('adult_label') else 'n/a'}), so there is nothing to compare.")

    # ---- Languages & countries
    with t_lang:
        l, r = st.columns(2)
        with l:
            if has("language_name"):
                show(pie(filtered["language_name"].value_counts(), "Original language share", n=8), 420)
        with r:
            if has("language_name"):
                ld = filtered["language_name"].value_counts().head(15).reset_index()
                ld.columns = ["Language", "Movies"]
                show(px.bar(ld, x="Language", y="Movies", title="Top original languages", color_discrete_sequence=[GOLD]), 420)
        l, r = st.columns(2)
        sp = explode(filtered, "spoken_language_list", "lang")
        with l:
            if not sp.empty:
                sc = sp["lang"].value_counts().head(15).reset_index()
                sc.columns = ["Language", "Movies"]
                show(px.bar(sc, x="Language", y="Movies", title="Top spoken languages", color_discrete_sequence=[TEAL]))
        with r:
            if not sp.empty:
                show(pie(sp["lang"].value_counts(), "Spoken language share", n=8))
        cd = explode(filtered, "countries_list", "country")
        if not cd.empty:
            l, r = st.columns(2)
            with l:
                cc = cd["country"].value_counts().head(15).reset_index()
                cc.columns = ["Country", "Movies"]
                fig = px.bar(cc, x="Movies", y="Country", orientation="h", title="Top production countries",
                             color_discrete_sequence=[ROSE])
                fig.update_layout(yaxis={"categoryorder": "total ascending"})
                show(fig, 440)
            with r:
                show(pie(cd["country"].value_counts(), "Production country share", n=8), 440)
            cnt = cd["country"].value_counts().reset_index()
            cnt.columns = ["Country", "Movies"]
            fig = px.choropleth(cnt, locations="Country", locationmode="country names", color="Movies",
                                color_continuous_scale="YlOrRd", title="Movies by production country")
            fig.update_geos(bgcolor="rgba(0,0,0,0)", showframe=False, showcountries=True, countrycolor="#3A3D4A",
                            landcolor="#161922", showocean=False, lakecolor="#0A0B10")
            show(fig, 480)

    # ---- Studios
    with t_studio:
        pc = explode(filtered, "production_list", "company")
        if pc.empty:
            st.info("No production company data in this selection.")
        else:
            l, r = st.columns(2)
            with l:
                cc = pc["company"].value_counts().head(15).reset_index()
                cc.columns = ["Company", "Movies"]
                fig = px.bar(cc, x="Movies", y="Company", orientation="h", title="Most active companies",
                             color_discrete_sequence=[GOLD])
                fig.update_layout(yaxis={"categoryorder": "total ascending"})
                show(fig, 460)
            with r:
                if has("revenue"):
                    cr = pc.groupby("company")["revenue"].sum().sort_values(ascending=False).head(15).reset_index()
                    cr.columns = ["Company", "Total revenue"]
                    fig = px.bar(cr, x="Total revenue", y="Company", orientation="h", title="Highest total revenue",
                                 color_discrete_sequence=[TEAL])
                    fig.update_layout(yaxis={"categoryorder": "total ascending"})
                    show(fig, 460)
            ra = rated(pc).groupby("company").agg(avg=("vote_average", "mean"), n=("title", "count"))
            ra = ra[ra["n"] >= 5].sort_values("avg", ascending=False).head(15).reset_index()
            if not ra.empty:
                fig = px.bar(ra, x="avg", y="company", orientation="h", title="Best rated companies (5+ movies)",
                             labels={"avg": "Average rating", "company": "Company"}, color_discrete_sequence=[SKY])
                fig.update_layout(yaxis={"categoryorder": "total ascending"})
                show(fig, 460)

    # ---- Correlations
    with t_corr:
        num = [c for c in ["budget", "revenue", "runtime", "vote_average", "vote_count", "popularity"] if has(c)]
        if len(num) >= 2:
            cm = filtered[num].corr().round(2)
            fig = px.imshow(cm, text_auto=True, color_continuous_scale="RdBu_r", zmin=-1, zmax=1,
                            title="Correlation between numeric columns")
            show(fig, 480)
            st.caption("1 means the two columns rise together, -1 means one rises as the other falls, 0 means no link.")


# ------------------------------------------------------------------
# PERFORMANCE
# ------------------------------------------------------------------
def page_performance():
    if not ready():
        return
    st.markdown("## Performance")
    y = filtered.dropna(subset=["release_year"]).copy()
    if not y.empty:
        y["release_year"] = y["release_year"].astype(int)
        yc = y.groupby("release_year").size().reset_index(name="Movies")
        fig = px.line(yc, x="release_year", y="Movies", markers=True, title="Release growth by year")
        fig.update_traces(line_color=GOLD, line_width=3)
        show(fig, 380)
        if len(yc) >= 2:
            ch = yc["Movies"].pct_change().iloc[-1] * 100
            if pd.notna(ch) and np.isfinite(ch):
                st.metric(f"Releases in {yc['release_year'].iloc[-1]} vs previous year", f"{ch:+.1f}%")
        if has("status"):
            ys = y.groupby(["release_year", "status"]).size().reset_index(name="Movies")
            show(px.bar(ys, x="release_year", y="Movies", color="status", title="Movies per year by status"), 380)
        l, r = st.columns(2)
        with l:
            rt = rated(y).groupby("release_year")["vote_average"].mean().reset_index()
            fig = px.line(rt, x="release_year", y="vote_average", markers=True, title="Average rating by year")
            fig.update_traces(line_color=TEAL, line_width=3)
            show(fig)
        with r:
            if has("revenue"):
                rv = y.groupby("release_year")["revenue"].sum().reset_index()
                fig = px.area(rv, x="release_year", y="revenue", title="Total revenue by year")
                fig.update_traces(line_color=ROSE, fillcolor="rgba(255,107,107,.20)")
                show(fig)
        l, r = st.columns(2)
        with l:
            if has("budget"):
                bd = y[y["budget"] > 0].groupby("release_year")["budget"].mean().reset_index()
                fig = px.bar(bd, x="release_year", y="budget", title="Average budget by year", color_discrete_sequence=[SKY])
                show(fig)
        with r:
            mo = y.groupby("release_month").size().reindex(MONTHS).reset_index(name="Movies")
            mo.columns = ["Month", "Movies"]
            show(px.bar(mo, x="Month", y="Movies", title="Releases by month", color_discrete_sequence=[GOLD]))
    else:
        st.info("Release dates are needed for yearly trends.")

    section("Budget and profitability",
            "Profit = revenue minus budget, using movies that have both a budget and a revenue above zero. Marketing costs are not included.")
    if has("budget") and has("revenue"):
        fin = filtered[(filtered["budget"] > 0) & (filtered["revenue"] > 0)].copy()
        if fin.empty:
            st.info("No movies in this selection have both a budget and a revenue above zero.")
            return
        fin["roi"] = fin["profit"] / fin["budget"] * 100
        c = st.columns(4)
        with c[0]: metric_card("Movies analysed", f"{len(fin):,}", "Budget and revenue above 0")
        with c[1]: metric_card("Average budget", money(fin["budget"].mean()), "Per movie")
        with c[2]: metric_card("Average revenue", money(fin["revenue"].mean()), "Per movie")
        with c[3]: metric_card("Average profit / loss", money(fin["profit"].mean()), "Revenue minus budget")
        fin["Result"] = np.select([fin["profit"] > 0, fin["profit"] < 0], ["Profit", "Loss"], "Break-even")
        st.write("")
        l, r = st.columns(2)
        with l:
            pc = fin["Result"].value_counts()
            fig = pie(pc, "Profit vs loss")
            fig.update_traces(marker=dict(colors=[{"Profit": TEAL, "Loss": ROSE, "Break-even": GOLD}[k] for k in pc.index],
                                          line=dict(color="#0A0B10", width=2)))
            show(fig)
        with r:
            sample = fin.sample(3000, random_state=1) if len(fin) > 3000 else fin
            fig = px.scatter(sample, x="budget", y="revenue", color="Result", opacity=0.7, hover_name="title",
                             log_x=True, log_y=True, title="Budget vs revenue (log scale)",
                             color_discrete_map={"Profit": TEAL, "Loss": ROSE, "Break-even": GOLD})
            show(fig)
        l, r = st.columns(2)
        with l:
            py = fin.dropna(subset=["release_year"]).groupby("release_year")["profit"].sum().reset_index()
            py["release_year"] = py["release_year"].astype(int)
            fig = px.bar(py, x="release_year", y="profit", title="Total profit by year", color="profit",
                         color_continuous_scale="RdYlGn")
            fig.update_layout(coloraxis_showscale=False)
            show(fig)
        with r:
            fg = explode(fin, "genres_list", "genre")
            if not fg.empty:
                gp = fg.groupby("genre")["profit"].sum().sort_values(ascending=False).reset_index()
                gp.columns = ["Genre", "Total profit"]
                show(px.bar(gp, x="Genre", y="Total profit", title="Total profit by genre", color_discrete_sequence=[GOLD]))
        bb = pd.cut(fin["budget"], [0, 1e5, 1e6, 1e7, 5e7, 1e8, 1e12],
                    labels=["Under $100K", "$100K – 1M", "$1M – 10M", "$10M – 50M", "$50M – 100M", "$100M+"])
        l, r = st.columns(2)
        with l:
            show(pie(bb.value_counts(), "Movies by budget size", n=10))
        with r:
            bg = fin.groupby(bb, observed=True)["roi"].median().reset_index()
            bg.columns = ["Budget size", "Median ROI %"]
            show(px.bar(bg, x="Budget size", y="Median ROI %", title="Median return on investment by budget size",
                        color_discrete_sequence=[TEAL]))

        section("Biggest hits and flops")
        a, b = st.columns(2)
        cols = ["title", "release_year", "budget", "revenue", "profit"]
        with a:
            st.caption("Highest profit")
            stretch(st.dataframe, fin.nlargest(10, "profit")[cols], hide_index=True)
        with b:
            st.caption("Biggest loss")
            stretch(st.dataframe, fin.nsmallest(10, "profit")[cols], hide_index=True)
        a, b = st.columns(2)
        with a:
            st.caption("Best return on investment (%)")
            stretch(st.dataframe, fin.nlargest(10, "roi")[cols + ["roi"]].round({"roi": 0}), hide_index=True)
        with b:
            st.caption("Highest revenue")
            stretch(st.dataframe, fin.nlargest(10, "revenue")[cols], hide_index=True)
    else:
        st.info("Budget and revenue columns are needed here.")


# ------------------------------------------------------------------
# DATA
# ------------------------------------------------------------------
def page_data():
    if not ready():
        return
    st.markdown("## Data")
    st.caption(f"Dataset: {source_label}  ·  {len(df):,} rows in total")
    order = ["movie_id", "title", "release_date", "status", "budget", "revenue", "runtime", "genres", "production",
             "countries", "spoken_language", "original_language", "vote_average", "vote_count", "popularity", "adult"]
    cols = [c for c in order if c in filtered.columns] + [c for c in ("overview", "poster") if c in filtered.columns]
    view = filtered[cols]
    q = st.text_input("Search the table by title", placeholder="Type part of a title")
    if q:
        view = view[view["title"].str.contains(q, case=False, na=False, regex=False)]
    st.caption(f"{len(view):,} movies shown")
    stretch(st.dataframe, view, hide_index=True, height=560)

    c1, c2, _ = st.columns([1, 1, 3])
    with c1:
        stretch(st.download_button, "Download CSV", view.to_csv(index=False).encode("utf-8-sig"),
                "cinescope_filtered_movies.csv", "text/csv")
    with c2:
        buf = io.BytesIO()
        out = view.copy()
        if "release_date" in out.columns:
            out["release_date"] = out["release_date"].dt.date
        out.to_excel(buf, index=False, engine="openpyxl")
        stretch(st.download_button, "Download Excel", buf.getvalue(), "cinescope_filtered_movies.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    section("Summary statistics")
    num = [c for c in ["budget", "revenue", "runtime", "vote_average", "vote_count", "popularity"] if c in filtered.columns]
    if num:
        stretch(st.dataframe, filtered[num].describe().round(2).T)
    section("Detected columns", "CineScope matched your file's columns to these fields.")
    st.write("*Found:* " + (", ".join(c for c in ALIASES if has(c)) or "none"))
    st.write("*Not found (related features are hidden):* " + (", ".join(c for c in ALIASES if not has(c)) or "none"))


# ------------------------------------------------------------------
# NAVIGATION
# ------------------------------------------------------------------
PAGE_HOME = st.Page(page_home, title="Home", icon="🏠", default=True)
PAGE_BROWSE = st.Page(page_browse, title="Browse", icon="🎞️", url_path="browse")
PAGE_MOVIE = st.Page(page_movie, title="Movie", icon="🎬", url_path="movie")
PAGE_ANALYTICS = st.Page(page_analytics, title="Analytics", icon="📊", url_path="analytics")
PAGE_PERFORMANCE = st.Page(page_performance, title="Performance", icon="📈", url_path="performance")
PAGE_DATA = st.Page(page_data, title="Data", icon="🗂️", url_path="data")

pages = [PAGE_HOME, PAGE_BROWSE, PAGE_MOVIE, PAGE_ANALYTICS, PAGE_PERFORMANCE, PAGE_DATA]
try:
    nav = st.navigation(pages, position="top")
except TypeError:
    nav = st.navigation(pages)
nav.run()

st.markdown('<div class="footer">CineScope · Built with Python, Pandas, Streamlit and Plotly</div>',
            unsafe_allow_html=True)