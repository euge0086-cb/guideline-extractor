"""
Guideline Reference Extractor
App Streamlit para extraer, enriquecer y clasificar referencias de guías clínicas.
"""

import streamlit as st
import tempfile
import os
import time
from io import BytesIO

# Importar pipeline
from pipeline import (
    extract_references_from_pdf,
    enrich_reference,
    classify_reference,
    export_to_excel,
)

# ─── Configuración de página ───────────────────────────────────────────────────
st.set_page_config(
    page_title="Guideline Reference Extractor",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─── CSS personalizado ─────────────────────────────────────────────────────────
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@300;400;600&family=IBM+Plex+Mono:wght@400&display=swap');

  html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif;
  }

  /* Header principal */
  .hero {
    background: linear-gradient(135deg, #0A2342 0%, #1B4F8A 100%);
    border-radius: 12px;
    padding: 2.5rem 2rem 2rem 2rem;
    margin-bottom: 2rem;
    color: white;
  }
  .hero h1 {
    font-size: 2rem;
    font-weight: 600;
    letter-spacing: -0.5px;
    margin: 0 0 0.4rem 0;
    color: white;
  }
  .hero p {
    font-size: 1rem;
    font-weight: 300;
    opacity: 0.85;
    margin: 0;
    color: white;
  }
  .hero .badge {
    display: inline-block;
    background: rgba(255,255,255,0.15);
    border: 1px solid rgba(255,255,255,0.3);
    border-radius: 20px;
    padding: 2px 12px;
    font-size: 0.75rem;
    font-family: 'IBM Plex Mono', monospace;
    margin-bottom: 1rem;
    color: #90CAF9;
  }

  /* Cards de estadísticas */
  .stat-card {
    background: #F8FAFD;
    border: 1px solid #E3EAF4;
    border-radius: 10px;
    padding: 1.2rem 1rem;
    text-align: center;
  }
  .stat-number {
    font-size: 2.2rem;
    font-weight: 600;
    color: #0A2342;
    line-height: 1;
  }
  .stat-label {
    font-size: 0.8rem;
    color: #6B7A99;
    margin-top: 4px;
    font-weight: 400;
  }

  /* Badges de tipo de estudio */
  .badge-rct { background:#E8F5E9; color:#2E7D32; border:1px solid #A5D6A7; border-radius:6px; padding:2px 8px; font-size:0.78rem; font-weight:600; }
  .badge-rct2 { background:#FFF8E1; color:#F57F17; border:1px solid #FFE082; border-radius:6px; padding:2px 8px; font-size:0.78rem; font-weight:600; }
  .badge-meta { background:#E3F2FD; color:#1565C0; border:1px solid #90CAF9; border-radius:6px; padding:2px 8px; font-size:0.78rem; font-weight:600; }
  .badge-reg { background:#FCE4EC; color:#880E4F; border:1px solid #F48FB1; border-radius:6px; padding:2px 8px; font-size:0.78rem; font-weight:600; }
  .badge-guide { background:#EDE7F6; color:#4527A0; border:1px solid #B39DDB; border-radius:6px; padding:2px 8px; font-size:0.78rem; font-weight:600; }
  .badge-other { background:#F5F5F5; color:#616161; border:1px solid #BDBDBD; border-radius:6px; padding:2px 8px; font-size:0.78rem; font-weight:600; }

  /* Tabla de referencias */
  .ref-row {
    background: white;
    border: 1px solid #E8EDF5;
    border-radius: 8px;
    padding: 1rem 1.2rem;
    margin-bottom: 0.5rem;
  }
  .ref-number {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.75rem;
    color: #9AA3B5;
    margin-bottom: 2px;
  }
  .ref-title {
    font-weight: 600;
    font-size: 0.9rem;
    color: #0A2342;
    margin-bottom: 3px;
  }
  .ref-meta {
    font-size: 0.8rem;
    color: #6B7A99;
  }
  .ref-links a {
    font-size: 0.78rem;
    color: #1565C0;
    text-decoration: none;
    margin-right: 12px;
  }
  .ref-links a:hover { text-decoration: underline; }

  /* Upload area */
  .upload-hint {
    font-size: 0.82rem;
    color: #6B7A99;
    margin-top: 0.5rem;
  }

  /* Progress */
  .progress-label {
    font-size: 0.85rem;
    color: #1B4F8A;
    font-weight: 600;
    margin-bottom: 4px;
  }

  /* Step indicator */
  .step {
    display: flex;
    align-items: center;
    gap: 10px;
    padding: 8px 0;
    font-size: 0.88rem;
    color: #6B7A99;
  }
  .step.done { color: #2E7D32; }
  .step.active { color: #1565C0; font-weight: 600; }
  .step-dot {
    width: 22px; height: 22px;
    border-radius: 50%;
    display: flex; align-items: center; justify-content: center;
    font-size: 0.7rem; font-weight: 700;
    flex-shrink: 0;
  }
  .step-dot.done { background:#E8F5E9; color:#2E7D32; }
  .step-dot.active { background:#E3F2FD; color:#1565C0; }
  .step-dot.pending { background:#F5F5F5; color:#BDBDBD; }

  /* Hide Streamlit branding */
  #MainMenu {visibility: hidden;}
  footer {visibility: hidden;}
  header {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# ─── Helpers ───────────────────────────────────────────────────────────────────

BADGE_MAP = {
    "RCT_primario":            '<span class="badge-rct">ECA primario</span>',
    "RCT_secundario":          '<span class="badge-rct2">ECA secundario</span>',
    "meta-analisis":           '<span class="badge-meta">Meta-análisis</span>',
    "registro_observacional":  '<span class="badge-reg">Registro/Cohorte</span>',
    "guia_clinica":            '<span class="badge-guide">Guía clínica</span>',
    "otro/no_clasificado":     '<span class="badge-other">Otro</span>',
}

COLOR_MAP = {
    "RCT_primario":           "#E8F5E9",
    "RCT_secundario":         "#FFF8E1",
    "meta-analisis":          "#E3F2FD",
    "registro_observacional": "#FCE4EC",
    "guia_clinica":           "#EDE7F6",
    "otro/no_clasificado":    "#F5F5F5",
}

def render_step(number, label, status):
    cls = status
    dot_content = "✓" if status == "done" else str(number)
    st.markdown(f"""
    <div class="step {cls}">
      <div class="step-dot {cls}">{dot_content}</div>
      {label}
    </div>
    """, unsafe_allow_html=True)

def type_counts(records):
    counts = {}
    for r in records:
        t = r.get("study_type_auto", "otro/no_clasificado")
        counts[t] = counts.get(t, 0) + 1
    return counts

def build_excel_bytes(records):
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp_path = tmp.name
    export_to_excel(records, tmp_path)
    with open(tmp_path, "rb") as f:
        data = f.read()
    os.unlink(tmp_path)
    return data


# ─── UI ───────────────────────────────────────────────────────────────────────

# Hero
st.markdown("""
<div class="hero">
  <div class="badge">v1.0 · Evidence-Based Medicine Tools</div>
  <h1>📚 Guideline Reference Extractor</h1>
  <p>Extrae, enriquece y clasifica todas las referencias de una guía clínica en PDF.<br>
  Obtén PMID, DOI, autores, año y tipo de estudio en un Excel listo para usar.</p>
</div>
""", unsafe_allow_html=True)

# Layout principal
col_left, col_right = st.columns([1, 2], gap="large")

with col_left:
    st.markdown("#### Subir guía clínica")
    uploaded_file = st.file_uploader(
        "Arrastra el PDF aquí",
        type=["pdf"],
        help="PDFs de guías ESC, ACC/AHA u otras. Funciona con layouts de 1 y 2 columnas.",
        label_visibility="collapsed"
    )
    st.markdown('<p class="upload-hint">Compatible con guías ESC, ACC/AHA y otras. Layouts de 1 y 2 columnas.</p>', unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("#### Proceso")

    # Estado inicial de los pasos
    if "records" not in st.session_state:
        render_step(1, "Extraer referencias del PDF", "pending")
        render_step(2, "Enriquecer con PubMed + CrossRef", "pending")
        render_step(3, "Clasificar por tipo de estudio", "pending")
        render_step(4, "Generar Excel descargable", "pending")

    st.markdown("---")
    st.markdown("#### ¿Qué detecta?")
    st.markdown("""
<div style="font-size:0.83rem; color:#444; line-height:1.8;">
<span class="badge-rct">ECA primario</span> Ensayo clínico aleatorizado<br>
<span class="badge-rct2">ECA secundario</span> Subanálisis / post-hoc<br>
<span class="badge-meta">Meta-análisis</span> Revisiones sistemáticas<br>
<span class="badge-reg">Registro/Cohorte</span> Estudios observacionales<br>
<span class="badge-guide">Guía clínica</span> Guidelines / consensos<br>
<span class="badge-other">Otro</span> No clasificado
</div>
""", unsafe_allow_html=True)

with col_right:
    if uploaded_file is None:
        # Estado vacío
        st.markdown("""
        <div style="background:#F8FAFD; border:2px dashed #C5D4E8; border-radius:12px;
                    padding:3rem 2rem; text-align:center; color:#6B7A99;">
          <div style="font-size:3rem; margin-bottom:1rem;">📄</div>
          <div style="font-size:1.1rem; font-weight:600; color:#0A2342; margin-bottom:0.5rem;">
            Sube un PDF para empezar
          </div>
          <div style="font-size:0.85rem;">
            El extractor detectará automáticamente la sección de referencias<br>
            y consultará PubMed y CrossRef para enriquecer los metadatos.
          </div>
        </div>
        """, unsafe_allow_html=True)

    else:
        # Botón de procesamiento
        if "records" not in st.session_state or st.session_state.get("last_file") != uploaded_file.name:
            if st.button("🔍 Extraer y enriquecer referencias", type="primary", use_container_width=True):
                with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp_pdf:
                    tmp_pdf.write(uploaded_file.read())
                    tmp_pdf_path = tmp_pdf.name

                # Panel de progreso
                progress_container = st.empty()
                with progress_container.container():
                    st.markdown('<p class="progress-label">Procesando...</p>', unsafe_allow_html=True)

                    step_area = st.empty()
                    progress_bar = st.progress(0)
                    status_text = st.empty()

                    # PASO 1: Extracción
                    with step_area.container():
                        render_step(1, "Extrayendo referencias del PDF...", "active")
                        render_step(2, "Enriquecer con PubMed + CrossRef", "pending")
                        render_step(3, "Clasificar por tipo de estudio", "pending")
                        render_step(4, "Generar Excel descargable", "pending")
                    progress_bar.progress(5)

                    raw_refs = extract_references_from_pdf(tmp_pdf_path)
                    n_refs = len(raw_refs)
                    progress_bar.progress(15)

                    with step_area.container():
                        render_step(1, f"Referencias extraídas: {n_refs}", "done")
                        render_step(2, "Enriqueciendo con PubMed + CrossRef...", "active")
                        render_step(3, "Clasificar por tipo de estudio", "pending")
                        render_step(4, "Generar Excel descargable", "pending")

                    # PASO 2+3: Enriquecimiento y clasificación
                    records = []
                    for i, ref_text in enumerate(raw_refs):
                        pct = 15 + int((i / n_refs) * 70)
                        progress_bar.progress(pct)
                        status_text.markdown(f'<p style="font-size:0.8rem;color:#6B7A99;">Referencia {i+1} de {n_refs}: {ref_text[:60]}...</p>', unsafe_allow_html=True)

                        record = enrich_reference(ref_text, i + 1)
                        record["study_type_auto"] = classify_reference(record)
                        records.append(record)

                    progress_bar.progress(88)
                    with step_area.container():
                        render_step(1, f"Referencias extraídas: {n_refs}", "done")
                        render_step(2, "Metadatos enriquecidos (PubMed + CrossRef)", "done")
                        render_step(3, "Clasificación automática completada", "done")
                        render_step(4, "Generando Excel...", "active")

                    # PASO 4: Excel
                    excel_bytes = build_excel_bytes(records)
                    progress_bar.progress(100)
                    status_text.empty()

                    with step_area.container():
                        render_step(1, f"Referencias extraídas: {n_refs}", "done")
                        render_step(2, "Metadatos enriquecidos (PubMed + CrossRef)", "done")
                        render_step(3, "Clasificación automática completada", "done")
                        render_step(4, "Excel generado", "done")

                # Guardar en session state
                st.session_state["records"] = records
                st.session_state["excel_bytes"] = excel_bytes
                st.session_state["last_file"] = uploaded_file.name
                os.unlink(tmp_pdf_path)
                progress_container.empty()
                st.rerun()

        # ── Resultados ─────────────────────────────────────────────────────────
        if "records" in st.session_state and st.session_state.get("last_file") == uploaded_file.name:
            records = st.session_state["records"]
            excel_bytes = st.session_state["excel_bytes"]
            counts = type_counts(records)

            # Botón de descarga destacado
            st.download_button(
                label="⬇️  Descargar Excel completo",
                data=excel_bytes,
                file_name=f"referencias_{uploaded_file.name.replace('.pdf','')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
            )

            st.markdown("---")

            # Estadísticas
            st.markdown("#### Resumen")
            stat_cols = st.columns(len(counts) if len(counts) <= 4 else 4)
            type_labels = {
                "RCT_primario": "ECAs primarios",
                "RCT_secundario": "ECAs secundarios",
                "meta-analisis": "Meta-análisis",
                "registro_observacional": "Registros",
                "guia_clinica": "Guías",
                "otro/no_clasificado": "Otros",
            }
            for i, (t, c) in enumerate(counts.items()):
                with stat_cols[i % 4]:
                    st.markdown(f"""
                    <div class="stat-card" style="border-top: 3px solid {COLOR_MAP.get(t,'#ccc')};">
                      <div class="stat-number">{c}</div>
                      <div class="stat-label">{type_labels.get(t, t)}</div>
                    </div>
                    """, unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)

            # Filtros
            st.markdown("#### Referencias")
            filter_col1, filter_col2 = st.columns([2, 1])
            with filter_col1:
                search_query = st.text_input("Buscar por título, autor o año...", placeholder="Ej: fibrinolysis 2019", label_visibility="collapsed")
            with filter_col2:
                type_options = ["Todos"] + list(counts.keys())
                type_filter = st.selectbox("Tipo", type_options, label_visibility="collapsed")

            # Filtrar registros
            filtered = records
            if type_filter != "Todos":
                filtered = [r for r in filtered if r.get("study_type_auto") == type_filter]
            if search_query:
                q = search_query.lower()
                filtered = [r for r in filtered if
                    q in (r.get("title") or "").lower() or
                    q in (r.get("authors") or "").lower() or
                    q in (r.get("year") or "").lower() or
                    q in (r.get("ref_raw") or "").lower()
                ]

            st.markdown(f'<p style="font-size:0.82rem;color:#9AA3B5;margin-bottom:0.5rem;">{len(filtered)} referencias</p>', unsafe_allow_html=True)

            # Lista de referencias
            for r in filtered:
                badge = BADGE_MAP.get(r.get("study_type_auto", ""), "")
                title = r.get("title") or r.get("ref_raw", "")[:100]
                authors = r.get("authors", "")
                year = r.get("year", "")
                journal = r.get("journal", "")
                pmid = r.get("pmid", "")
                doi = r.get("doi", "")
                meta = " · ".join(filter(None, [authors[:60] + ("..." if len(authors) > 60 else ""), year, journal[:40]]))

                links = ""
                if pmid:
                    links += f'<a href="https://pubmed.ncbi.nlm.nih.gov/{pmid}/" target="_blank">PubMed →</a>'
                if doi:
                    links += f'<a href="https://doi.org/{doi}" target="_blank">DOI →</a>'

                st.markdown(f"""
                <div class="ref-row">
                  <div class="ref-number">#{r.get('ref_number', '')} &nbsp; {badge}</div>
                  <div class="ref-title">{title}</div>
                  <div class="ref-meta">{meta}</div>
                  {"<div class='ref-links'>" + links + "</div>" if links else ""}
                </div>
                """, unsafe_allow_html=True)

            if st.button("🔄 Procesar otro PDF", use_container_width=False):
                for key in ["records", "excel_bytes", "last_file"]:
                    if key in st.session_state:
                        del st.session_state[key]
                st.rerun()
