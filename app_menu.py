"""
╔══════════════════════════════════════════════════════════════════╗
║         NOIR MENU STUDIO  ·  v8.0  ·  Premium Edition        ║
╠══════════════════════════════════════════════════════════════════╣
║  Release Notes v8.0:                                             ║
║  • UI Dark "Noir": Antracite (#111), Oro (#d4a) e Grigio Seta    ║
║  • Swiss Design Engine: Tipografia e spaziature millimetriche    ║
║  • Domino Page Builder: Propagazione intelligente dei salti      ║
║  • State-Machine Parsers: Estrazione Word ultra-robusta          ║
║  • Architettura "Boxed": Codice pulito, modulare e manutenibile  ║
╚══════════════════════════════════════════════════════════════════╝
"""

import streamlit as st
import docx
import re
import json
import base64
import io
import unicodedata
import pandas as pd

# Tentativo di import WeasyPrint per PDF
try:
    from weasyprint import HTML as WeasyHTML
    PDF_DISPONIBILE = True
except Exception:
    PDF_DISPONIBILE = False

# Configurazione Iniziale Streamlit
st.set_page_config(
    page_title="Noir Menu Studio",
    layout="wide",
    page_icon="🍽️",
    initial_sidebar_state="expanded"
)


# ═══════════════════════════════════════════════════════════════════
# 1. CONFIG & COSTANTI
# ═══════════════════════════════════════════════════════════════════

TEMPLATE_CLASSICO = "🥂 Classico"
TEMPLATE_MODERNO  = "◼ Moderno"
TEMPLATE_RUSTICO  = "🌿 Rustico"
TEMPLATE_OPTIONS  = [TEMPLATE_CLASSICO, TEMPLATE_MODERNO, TEMPLATE_RUSTICO]

# Palette Noir Studio Interface
COLOR_BG      = "#111111"
COLOR_SIDEBAR = "#1a1a1a"
COLOR_ACCENT  = "#d4a843" # Oro opaco
COLOR_TEXT    = "#e0e0e0" # Grigio seta
COLOR_SUBTEXT = "#aaaaaa"

# Configurazione Template Menu (Output Grafico - Swiss Design Precision)
TEMPLATE_CSS = {
    TEMPLATE_CLASSICO: {
        'font_import': "@import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,300;0,400;0,500;0,600;0,700;1,400&display=swap');",
        'font_family': "'Cormorant Garamond', serif",
        'bg_color':    '#fdfaf6',
        'titolo_color':'#1a1a1a',
        'cat_color':   '#8a6d3b',
        'nome_weight': '600',
        'desc_it_color':'#444444',
        'desc_en_color':'#999999',
        'prezzo_color': '#1a1a1a',
        'sep_color':   '#e0d7c6',
        'footer_color':'#aaaaaa',
        'letter_spacing': '0.01em',
        'cat_spacing': '0.25em',
        'line_height': '1.6',
        'text_transform': 'none',
    },
    TEMPLATE_MODERNO: {
        'font_import': "@import url('https://fonts.googleapis.com/css2?family=Inter:wght@200;300;400;600;700&display=swap');",
        'font_family': "'Inter', sans-serif",
        'bg_color':    '#ffffff',
        'titolo_color':'#000000',
        'cat_color':   '#000000',
        'nome_weight': '600',
        'desc_it_color':'#333333',
        'desc_en_color':'#999999',
        'prezzo_color': '#000000',
        'sep_color':   '#f0f0f0',
        'footer_color':'#bbbbbb',
        'letter_spacing': '-0.02em',
        'cat_spacing': '0.12em',
        'line_height': '1.4',
        'text_transform': 'uppercase',
    },
    TEMPLATE_RUSTICO: {
        'font_import': "@import url('https://fonts.googleapis.com/css2?family=Playfair+Display:ital,wght@0,400;0,700;0,900;1,400&display=swap');",
        'font_family': "'Playfair Display', serif",
        'bg_color':    '#f9f4ee',
        'titolo_color':'#2c1810',
        'cat_color':   '#5d4037',
        'nome_weight': '800',
        'desc_it_color':'#4e342e',
        'desc_en_color':'#8d6e63',
        'prezzo_color': '#2c1810',
        'sep_color':   '#d7ccc8',
        'footer_color':'#a1887f',
        'letter_spacing': '0em',
        'cat_spacing': '0.15em',
        'line_height': '1.5',
        'text_transform': 'none',
    },
}


# ═══════════════════════════════════════════════════════════════════
# 2. PREMIUM UI (STUDIO INTERFACE)
# ═══════════════════════════════════════════════════════════════════

def apply_premium_ui():
    st.markdown(f"""
    <style>
    /* Global Background and Text */
    .stApp {{
        background-color: {COLOR_BG};
        color: {COLOR_TEXT};
        font-family: 'Inter', sans-serif;
    }}

    /* Sidebar Styling */
    [data-testid="stSidebar"] {{
        background-color: {COLOR_SIDEBAR};
        border-right: 1px solid #222;
    }}
    [data-testid="stSidebar"] .stMarkdown h2 {{
        color: {COLOR_ACCENT} !important;
        font-family: 'Playfair Display', serif;
        font-weight: 700;
        letter-spacing: 0.05em;
        margin-top: 1rem;
    }}

    /* Elegant Buttons Noir */
    .stButton > button {{
        background-color: transparent;
        color: {COLOR_ACCENT};
        border: 1px solid {COLOR_ACCENT};
        border-radius: 2px;
        transition: all 0.4s cubic-bezier(0.165, 0.84, 0.44, 1);
        text-transform: uppercase;
        letter-spacing: 0.15em;
        font-weight: 500;
        font-size: 0.75rem;
        padding: 0.6rem 1rem;
        width: 100%;
    }}
    .stButton > button:hover {{
        background-color: {COLOR_ACCENT};
        color: #000;
        border-color: {COLOR_ACCENT};
        box-shadow: 0 4px 12px rgba(212, 168, 67, 0.2);
    }}
    div.stButton > button[kind="primary"] {{
        background-color: {COLOR_ACCENT};
        color: #000;
        border: none;
    }}
    div.stButton > button[kind="primary"]:hover {{
        background-color: #e5bc5e;
        box-shadow: 0 4px 15px rgba(212, 168, 67, 0.4);
    }}

    /* Custom Tabs */
    .stTabs [data-baseweb="tab-list"] {{
        gap: 32px;
        background-color: transparent;
        border-bottom: 1px solid #222;
    }}
    .stTabs [data-baseweb="tab"] {{
        color: {COLOR_SUBTEXT};
        border-bottom: 2px solid transparent;
        transition: all 0.3s;
        padding: 12px 0;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.1em;
        font-size: 0.8rem;
    }}
    .stTabs [data-baseweb="tab"]:hover {{
        color: {COLOR_ACCENT};
    }}
    .stTabs [data-baseweb="tab"][aria-selected="true"] {{
        color: {COLOR_ACCENT} !important;
        border-bottom-color: {COLOR_ACCENT} !important;
    }}

    /* Expanders & Widgets */
    div[data-testid="stExpander"] {{
        border: 1px solid #222 !important;
        background-color: #151515 !important;
        border-radius: 4px !important;
    }}
    .streamlit-expanderHeader {{
        color: {COLOR_SUBTEXT} !important;
        font-weight: 600 !important;
    }}
    .stTextInput input, .stSelectbox [data-baseweb="select"], .stTextArea textarea {{
        background-color: #0a0a0a !important;
        border: 1px solid #222 !important;
        color: {COLOR_TEXT} !important;
        border-radius: 2px !important;
    }}

    /* Page Builder Cards */
    .builder-card {{
        background: #181818;
        border: 1px solid #282828;
        border-radius: 4px;
        padding: 14px;
        margin-bottom: 10px;
        border-left: 2px solid {COLOR_ACCENT};
        box-shadow: 0 4px 6px rgba(0,0,0,0.3);
        transition: transform 0.2s;
    }}
    .builder-card:hover {{
        transform: translateY(-2px);
        border-color: #333;
    }}

    /* Elegant Banner Noir */
    .menu-banner {{
        background: #000;
        padding: 40px 40px;
        border-radius: 4px;
        margin-bottom: 40px;
        border-bottom: 1px solid {COLOR_ACCENT}44;
        position: relative;
        overflow: hidden;
    }}
    .menu-banner::before {{
        content: "";
        position: absolute;
        top: 0; left: 0; width: 4px; height: 100%;
        background-color: {COLOR_ACCENT};
    }}
    .menu-banner h1 {{
        color: {COLOR_ACCENT};
        font-family: 'Playfair Display', serif;
        margin: 0;
        font-size: 2.8em;
        font-weight: 700;
        letter-spacing: -0.02em;
    }}
    .menu-banner p {{
        color: {COLOR_SUBTEXT};
        text-transform: uppercase;
        letter-spacing: 0.3em;
        font-size: 0.75em;
        margin-top: 10px;
        font-weight: 500;
    }}

    /* Split-Screen Desktop Experience */
    @media (min-width: 1200px) {{
        [data-testid="stVerticalBlock"] > [data-testid="stColumn"]:nth-child(1) {{
            height: calc(100vh - 4rem);
            overflow-y: auto;
            padding-right: 30px;
            scrollbar-width: thin;
            scrollbar-color: #333 transparent;
        }}
        [data-testid="stVerticalBlock"] > [data-testid="stColumn"]:nth-child(2) [data-testid="stVerticalBlock"] {{
            position: sticky;
            top: 1rem;
            height: calc(100vh - 3rem);
            overflow-y: auto;
            border-left: 1px solid #222;
            padding-left: 35px;
            background: #050505;
            border-radius: 4px;
        }}
    }}

    /* Slider and Metrics */
    [data-testid="stMetricValue"] {{
        color: {COLOR_ACCENT} !important;
        font-family: 'Playfair Display', serif;
        font-size: 1.8rem !important;
    }}
    .stSlider [data-baseweb="slider"] [role="slider"] {{
        background-color: {COLOR_ACCENT};
    }}
    </style>
    """, unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════
# 3. UTILS
# ═══════════════════════════════════════════════════════════════════

def get_image_base64(uploaded_file):
    if uploaded_file is not None:
        return f"data:{uploaded_file.type};base64,{base64.b64encode(uploaded_file.getvalue()).decode()}"
    return None

def _safe_str(val) -> str:
    if val is None: return ""
    try:
        if pd.isna(val): return ""
    except: pass
    return str(val).strip() if str(val).lower() != "nan" else ""

def _safe_int(val, default=1) -> int:
    try: return max(1, int(float(str(val))))
    except: return default

def _split_bilingue(testo: str) -> tuple:
    if ' / ' in testo:
        p = testo.split(' / ', 1)
        return p[0].strip(), p[1].strip()
    return testo.strip(), ''


# ═══════════════════════════════════════════════════════════════════
# 4. PARSERS (WORD EXTRACTION)
# ═══════════════════════════════════════════════════════════════════

_RE_EURO = re.compile(r'€')
_RE_PREZZO_FULL = re.compile(r'(€\s*\d[\d.,]*(?:\s*/\s*(?:per\s*100\s*g|l\'etto|etto|kg))?|\d[\d.,]*\s*€)', re.IGNORECASE)
_RE_PREZZO_SOLO = re.compile(r'^€\s*[\d.,]+\s*$')
_RE_ALLERG_TONDE = re.compile(r'\((?:Allergen[^\)]*:|)\s*[\d,\s]+\)', re.IGNORECASE)
_RE_ALLERG_QUADRE = re.compile(r'\[\s*[\d,\s]+\s*\]')

def _normalizza_prezzo(raw: str) -> str:
    num = re.sub(r'[€\s]', '', raw).replace(',', '.')
    return f"€ {num}"

def _estrai_allergeni_auto(testo: str):
    trovati_q = [m.group(0).strip('[] ').strip() for m in _RE_ALLERG_QUADRE.finditer(testo)]
    t_no_q = _RE_ALLERG_QUADRE.sub('', testo).strip()
    trovati_t = []
    for m in _RE_ALLERG_TONDE.finditer(t_no_q):
        pulito = re.sub(r'(Allergen[^\)]*:|Allergens\s*:)', '', m.group(0), flags=re.IGNORECASE).strip('() ').strip()
        trovati_t.append(pulito)
    clean = _RE_ALLERG_TONDE.sub('', t_no_q).strip().strip('*').strip()
    all_found = list(set(trovati_q + trovati_t))
    return clean, ', '.join(all_found)

def parser_smart(doc) -> list:
    """
    Parser universale potenziato con macchina a stati.
    Gestisce formati eterogenei, prezzi su righe separate e bilinguismo.
    """
    piatti = []
    cat_it, cat_en = "Menu", ""
    piatto = None
    desc_count = 0

    # Pre-analisi per pulizia
    testi_puliti = []
    for p in doc.paragraphs:
        t = p.text.strip()
        if not t: continue
        # Rilevamento grassetto (spesso usato per i nomi)
        is_bold = any(r.bold for r in p.runs if r.text.strip())
        # Rilevamento corsivo (spesso usato per le descrizioni)
        is_italic = all((r.italic or not r.text.strip()) for r in p.runs)
        testi_puliti.append({'t': t, 'bold': is_bold, 'italic': is_italic})

    for entry in testi_puliti:
        t = entry['t']
        ha_euro = bool(_RE_EURO.search(t))
        prezzo_solo = bool(_RE_PREZZO_SOLO.match(t))

        # 1. Rilevamento Categoria (Tutto Maiuscolo o molto corto e Bold)
        is_cat = (t == t.upper() and len(t) > 3) or (not ha_euro and len(t.split()) <= 3 and entry['bold'])
        if is_cat:
            if piatto: piatti.append(piatto); piatto = None
            cat_it, cat_en = _split_bilingue(t.strip('*'))
            desc_count = 0
            continue

        # 2. Rilevamento Prezzo Isolato (chiude il piatto precedente se esiste)
        if prezzo_solo and piatto:
            piatto['Prezzo'] = _normalizza_prezzo(t)
            piatti.append(piatto)
            piatto = None
            continue

        # 3. Rilevamento Nuovo Piatto (Bold o riga con Prezzo Inline)
        if entry['bold'] or (ha_euro and not prezzo_solo):
            if piatto: piatti.append(piatto)

            t_c, al = _estrai_allergeni_auto(t)
            m_p = _RE_PREZZO_FULL.search(t_c)

            if m_p:
                prz = _normalizza_prezzo(m_p.group(1))
                nome_raw = t_c[:m_p.start()].strip().strip('* —-.')
                desc_extra = t_c[m_p.end():].strip()
            else:
                prz = ""
                nome_raw = t_c.strip('*')
                desc_extra = ""

            n_it, n_en = _split_bilingue(nome_raw)
            piatto = {
                'Categoria IT': cat_it, 'Categoria EN': cat_en,
                'Nome IT': n_it, 'Nome EN': n_en,
                'Descrizione IT': desc_extra, 'Descrizione EN': '',
                'Prezzo': prz, 'Allergeni': al
            }
            desc_count = 0

            # Se la riga conteneva già tutto (nome e prezzo), e non è bold,
            # potremmo volerlo salvare subito, ma aspettiamo la riga successiva per eventuali descrizioni
            continue

        # 4. Rilevamento Descrizioni (testo normale o italic sotto un piatto)
        if piatto:
            t_c, al = _estrai_allergeni_auto(t)
            if al:
                piatto['Allergeni'] = (piatto['Allergeni'] + ', ' + al).strip(', ')

            if t_c:
                # Se è la prima riga dopo il nome, va in IT, altrimenti EN
                if desc_count == 0 and not piatto['Descrizione IT']:
                    piatto['Descrizione IT'] = t_c
                elif desc_count == 0 and piatto['Descrizione IT']:
                    piatto['Descrizione EN'] = t_c
                    desc_count = 1
                else:
                    piatto['Descrizione EN'] = (piatto['Descrizione EN'] + ' ' + t_c).strip()
                desc_count += 1

    if piatto: piatti.append(piatto)

    # Post-process: pulizia spazi doppi
    for p in piatti:
        p['Descrizione IT'] = re.sub(r'\s{2,}', ' ', p['Descrizione IT']).strip()
        p['Descrizione EN'] = re.sub(r'\s{2,}', ' ', p['Descrizione EN']).strip()

    return piatti


# ═══════════════════════════════════════════════════════════════════
# 5. RENDERING ENGINE (SWISS DESIGN PRECISION)
# ═══════════════════════════════════════════════════════════════════

def genera_html(df, logo_b64, bg_b64, titolo, footer, l_size, template_key) -> str:
    """
    Genera l'HTML del menu applicando i principi dello Swiss Design:
    precisione tipografica, ampi spazi bianchi e gerarchia visiva chiara.
    """
    t = TEMPLATE_CSS.get(template_key, TEMPLATE_CSS[TEMPLATE_CLASSICO])
    df_v = df[df['Visibile'] == True].sort_values(['Pagina', 'Ordine']).reset_index(drop=True)

    html = f"""
    <!DOCTYPE html><html><head><meta charset="UTF-8"><style>
    {t['font_import']}
    @page {{ size: A4; margin: 0; }}
    body {{
        margin: 0; padding: 0; background: #222;
        font-family: {t['font_family']}; -webkit-print-color-adjust: exact;
    }}
    .foglio {{
        width: 210mm; height: 297mm;
        background: {f"url('{bg_b64}') center/cover" if bg_b64 else t['bg_color']};
        position: relative; margin: 30px auto;
        box-shadow: 0 20px 60px rgba(0,0,0,0.8);
        page-break-after: always; display: flex;
        flex-direction: column; overflow: hidden;
        color: {t['titolo_color']};
    }}
    .content {{
        padding: 28mm 32mm; flex-grow: 1;
        text-align: center; display: flex;
        flex-direction: column; align-items: center;
    }}
    .header-area {{ margin-bottom: 12mm; flex-shrink: 0; }}
    .titolo-area {{ margin-bottom: 15mm; max-width: 90%; }}
    .titolo {{
        font-size: 3.4em; font-weight: 300; color: {t['titolo_color']};
        letter-spacing: 0.18em; text-transform: uppercase; line-height: 1.1;
    }}
    .piatti-container {{ width: 100%; }}
    .cat {{
        font-size: 1.45em; color: {t['cat_color']};
        text-transform: uppercase; letter-spacing: {t['cat_spacing']};
        margin: 14mm 0 10mm 0; border-bottom: 0.5px solid {t['sep_color']};
        display: inline-block; padding: 0 30px 8px 30px;
        font-weight: 500;
    }}
    .piatto {{
        margin-bottom: 9mm; line-height: {t['line_height']};
        page-break-inside: avoid; width: 100%;
    }}
    .n-p {{
        display: flex; justify-content: center; align-items: baseline;
        gap: 16px; margin-bottom: 4px;
    }}
    .nome {{
        font-size: 1.35em; font-weight: {t['nome_weight']};
        color: {t['titolo_color']}; letter-spacing: {t['letter_spacing']};
        text-transform: {t['text_transform']};
    }}
    .prezzo {{ font-weight: 600; color: {t['prezzo_color']}; font-size: 1.15em; }}
    .d-it {{
        font-size: 0.98em; font-style: italic; color: {t['desc_it_color']};
        max-width: 85%; margin: 6px auto; line-height: 1.6;
    }}
    .d-en {{
        font-size: 0.88em; font-style: italic; color: {t['desc_en_color']};
        opacity: 0.65; max-width: 85%; margin: 3px auto;
    }}
    .foot {{
        position: absolute; bottom: 15mm; width: 100%; text-align: center;
        font-size: 0.72em; color: {t['footer_color']}; padding: 0 32mm;
        box-sizing: border-box; text-transform: uppercase;
        letter-spacing: 0.15em; opacity: 0.8;
    }}
    @media print {{
        body {{ background: none; }}
        .foglio {{ margin: 0; box-shadow: none; }}
    }}
    </style></head><body>
    """

    curr_p = -1; curr_c = None
    for _, row in df_v.iterrows():
        p = _safe_int(row['Pagina'])
        if p != curr_p:
            if curr_p != -1:
                html += f"</div></div><div class='foot'>{footer.replace(chr(10),'<br>')}</div></div>"
            html += f"<div class='foglio'><div class='content'>"

            if p == 1:
                html += "<div class='header-area'>"
                if logo_b64:
                    html += f"<img src='{logo_b64}' style='max-width: {l_size}px; max-height: 45mm; object-fit: contain;'>"
                html += "</div>"
                if titolo:
                    html += f"<div class='titolo-area'><div class='titolo'>{titolo}</div></div>"

            html += "<div class='piatti-container'>"
            curr_p = p; curr_c = None

        c = _safe_str(row['Categoria IT'])
        if c != curr_c:
            col_cat = _safe_str(row.get('Colore Cat.', t['cat_color']))
            html += f"<div class='cat' style='color:{col_cat}; border-bottom-color:{col_cat}33;'>{c}</div>"
            curr_c = c

        sc = float(row.get('Scala Piatto', 1.0)); ex = float(row.get('Spazio Extra', 0.0))
        html += f"<div class='piatto' style='transform: scale({sc}); margin-bottom: {9+ex*6}mm;'>"
        html += f"<div class='n-p'><span class='nome'>{row['Nome IT']}</span>"
        if row['Prezzo']:
            html += f"<span class='prezzo'>{row['Prezzo']}</span>"
        html += "</div>"
        if row['Descrizione IT']:
            html += f"<div class='d-it'>{row['Descrizione IT']}</div>"
        if row['Descrizione EN']:
            html += f"<div class='d-en'>{row['Descrizione EN']}</div>"
        html += "</div>"

    html += f"</div></div><div class='foot'>{footer.replace(chr(10),'<br>')}</div></div></body></html>"
    return html


# ═══════════════════════════════════════════════════════════════════
# 6. DATA & DOMINO BUILDER
# ═══════════════════════════════════════════════════════════════════

def _assicura_colonne(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty: return pd.DataFrame()
    cols = {'Ordine': lambda n: list(range(1, n+1)), 'Pagina': 1, 'Visibile': True, 'Scala Piatto': 1.0, 'Spazio Extra': 0.0, 'Colore Cat.': '#8a6d3b'}
    for c, d in cols.items():
        if c not in df.columns: df[c] = d(len(df)) if callable(d) else d
    df['Pagina'] = df['Pagina'].apply(_safe_int)
    df['Ordine'] = pd.to_numeric(df['Ordine'], errors='coerce').fillna(0).astype(int)
    return df

def apply_domino(idx, new_pag):
    df = st.session_state.dati_menu.copy()
    old_pag = df.at[idx, 'Pagina']
    df.at[idx, 'Pagina'] = new_pag
    if new_pag > old_pag:
        for i in range(idx + 1, len(df)):
            if df.at[i, 'Pagina'] < new_pag: df.at[i, 'Pagina'] = new_pag
    st.session_state.dati_menu = df

def render_builder():
    st.markdown(f"### <span style='color:{COLOR_ACCENT}'>🏗️ Visual Domino Builder</span>", unsafe_allow_html=True)
    st.caption("L'effetto Domino sposta automaticamente tutti i piatti successivi per mantenere l'ordine logico del menu.")

    df = st.session_state.dati_menu
    if df.empty:
        return st.info("Importa dati nel Database per attivare il builder.")

    # Reset index per stabilità
    df = df.reset_index(drop=True)
    st.session_state.dati_menu = df

    max_p = int(df['Pagina'].max())
    num_pages_to_show = max_p + 1

    # Grid di pagine (3 per riga)
    for row_idx in range(0, num_pages_to_show, 3):
        cols = st.columns(3)
        for col_idx in range(3):
            page_num = row_idx + col_idx + 1
            if page_num > num_pages_to_show:
                break

            with cols[col_idx]:
                st.markdown(f"#### 📄 Pagina {page_num}")
                df_p = df[df['Pagina'] == page_num].sort_values('Ordine')

                if df_p.empty:
                    st.markdown("<div style='color:#444; font-style:italic; padding:15px; border:1px dashed #333; border-radius:4px; text-align:center; margin-bottom:20px;'>Nessun piatto</div>", unsafe_allow_html=True)
                else:
                    for idx_in_df, row in df_p.iterrows():
                        with st.container():
                            # Card con stili Premium Noir
                            st.markdown(f"""
                            <div class="builder-card">
                                <div style="color:{COLOR_ACCENT}; font-weight:700; margin-bottom:4px;">{row['Nome IT']}</div>
                                <div style="color:{COLOR_SUBTEXT}; font-size:0.75rem; text-transform:uppercase; letter-spacing:0.05em;">{row['Categoria IT']}</div>
                            </div>
                            """, unsafe_allow_html=True)

                            # Input numerico per cambio pagina (Domino)
                            new_v = st.number_input(
                                "Pag.", 1, 30, int(row['Pagina']),
                                key=f"dom_{idx_in_df}_{page_num}",
                                label_visibility="collapsed"
                            )

                            if new_v != int(row['Pagina']):
                                apply_domino(idx_in_df, new_v)
                                st.rerun()


# ═══════════════════════════════════════════════════════════════════
# 7. MAIN APP
# ═══════════════════════════════════════════════════════════════════

def main():
    apply_premium_ui()
    if 'dati_menu' not in st.session_state: st.session_state.dati_menu = pd.DataFrame()
    st.markdown('<div class="menu-banner"><h1>NOIR MENU STUDIO</h1><p>Premium Editorial Design Studio</p></div>', unsafe_allow_html=True)

    with st.sidebar:
        st.markdown("## 🎨 CONFIGURAZIONE")
        with st.expander("DESIGN & TEMPLATE", expanded=True):
            template = st.selectbox("Stile", TEMPLATE_OPTIONS)
            titolo = st.text_input("Titolo Menu", "Specialità")
            footer = st.text_area("Piè di pagina", "Servizio incluso.")
            logo = st.file_uploader("Logo", type=['png','jpg'])
            l_size = st.slider("Logo Size", 50, 400, 150)
            bg = st.file_uploader("Texture Sfondo", type=['png','jpg'])
        if not st.session_state.dati_menu.empty:
            with st.expander("ESPORTAZIONE", expanded=False):
                st.download_button("💾 Progetto JSON", st.session_state.dati_menu.to_json(), "menu.json")
                if PDF_DISPONIBILE and st.button("📄 Genera PDF"):
                    h = genera_html(st.session_state.dati_menu, get_image_base64(logo), get_image_base64(bg), titolo, footer, l_size, template)
                    pdf = WeasyHTML(string=h).write_pdf(); st.download_button("⬇️ Scarica PDF", pdf, "menu.pdf", "application/pdf")

    tab_edit, tab_build = st.tabs(["🗄️ DATABASE PIATTI", "🏗️ PAGE BUILDER"])
    col_ctrl, col_prev = st.columns([1.2, 1], gap="large")

    with col_ctrl:
        # Semplificazione: mostriamo l'editor nel tab 1 e il builder nel tab 2
        with tab_edit:
            st.markdown(f"### <span style='color:{COLOR_ACCENT}'>📂 Caricamento Word</span>", unsafe_allow_html=True)
            f_it = st.file_uploader("Documento .docx", type=['docx'])
            if st.button("✨ ANALIZZA E IMPORTA", type="primary"):
                if f_it:
                    doc = docx.Document(f_it)
                    st.session_state.dati_menu = _assicura_colonne(pd.DataFrame(parser_smart(doc)))
                    st.rerun()
            if not st.session_state.dati_menu.empty:
                st.markdown("---")
                edited = st.data_editor(st.session_state.dati_menu, use_container_width=True, num_rows="dynamic")
                if st.button("💾 SALVA MODIFICHE"): st.session_state.dati_menu = edited; st.rerun()

        with tab_build:
            render_builder()

    with col_prev:
        st.markdown(f"### <span style='color:{COLOR_ACCENT}'>👁️ ANTEPRIMA LIVE</span>", unsafe_allow_html=True)
        if not st.session_state.dati_menu.empty:
            html_l = genera_html(st.session_state.dati_menu, get_image_base64(logo), get_image_base64(bg), titolo, footer, l_size, template)
            st.components.v1.html(html_l, height=1000, scrolling=True)
        else: st.info("In attesa di dati per l'anteprima.")

if __name__ == "__main__":
    main()
