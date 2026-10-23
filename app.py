# -*- coding: utf-8 -*-
"""
Generador de descripciones de producto — Kbas Office
Sube fotos (se agrupan por referencia) + el Excel de datos, elige el prompt,
y la herramienta saca, por cada producto: descripción + 3 titulares (a Claude).
Descarga un Excel: Ref -> Tipo / Color / Medidas / Descripción / Titulares.
"""

import io
import os
import re
import base64
import zipfile
import pandas as pd
import streamlit as st

# ---------------- Config ----------------
MODELO_CLAUDE = "claude-sonnet-5-5"     # se puede cambiar por claude-haiku-5-5 (más barato) o claude-opus-5-5 (mejor)
SUFIJOS_VISTA = ("_LT", "_D", "_H")      # lateral-trasera, detalle, con asa  (la principal no lleva sufijo)
MAX_FOTOS = 4                            # principal + _D + _LT + _H

# Prompts guardados (editables en la interfaz)
PROMPT_ECI = (
    "Quiero que actúes como un experto en SEO y comunicación online persuasiva. "
    "Me gustaría para estas imágenes de un mismo bolso una descripción comercial que tenga "
    "buenas prácticas SEO en tono sofisticado pero además un copy elaborado y creativo. "
    "Que sea una versión para ficha producto, por tanto no muy extensa, siempre con enfoque "
    "SEO y copy persuasivo. Y también proponme tres Titulares cortos que incluyan el tipo de "
    "producto, el (tamaño siempre que se conozca), el material (siempre que sea destacable) y "
    "una característica relevante pero, no debe poner el color.\n\n"
    "Ejemplo de estilo:\n"
    "Titular: Bolso bowling tamaño grande en tejido estructurado con asas acolchadas\n"
    "Descripción: Este bolso tipo bowling reinventa el estilo casual con un enfoque sofisticado y "
    "contemporáneo. Confeccionado en tejido estructurado de alta resistencia color kaki, destaca por "
    "su diseño fruncido que aporta volumen y personalidad, mientras que las asas acolchadas elevan la "
    "comodidad y el carácter del conjunto. Su formato amplio lo convierte en el aliado perfecto para el "
    "día a día, ofreciendo espacio y organización sin renunciar al estilo. Incorpora cierre superior con "
    "cremallera y una práctica correa ajustable para llevarlo también al hombro o cruzado, adaptándose a "
    "tu ritmo con total versatilidad. Un bolso urbano con esencia elegante, ideal para quienes buscan "
    "funcionalidad, tendencia y un diseño distintivo en un solo accesorio."
)
PROMPT_KBAS = (
    "Actúa como copywriter de moda para la web propia de Kbas. Escribe una descripción de ficha de "
    "producto, cercana y con estilo, con buen SEO y sin ser muy extensa. Añade tres titulares cortos "
    "que incluyan tipo de producto, tamaño (si se conoce), material (si es destacable) y una "
    "característica relevante; los titulares no deben incluir el color."
)
PROMPTS_GUARDADOS = {
    "El Corte Inglés": PROMPT_ECI,
    "Kbas (web propia)": PROMPT_KBAS,
}

# Formato de salida que se añade SIEMPRE al prompt para poder separar en columnas
INSTRUCCION_FORMATO = (
    "\n\n---\nIMPORTANTE: No incluyas las medidas exactas (cm) en la descripción; van en otro campo. "
    "Responde EXACTAMENTE en este formato, sin ningún texto adicional antes ni después:\n"
    "DESCRIPCIÓN: <la descripción>\n"
    "TITULAR 1: <titular 1>\n"
    "TITULAR 2: <titular 2>\n"
    "TITULAR 3: <titular 3>"
)

# ---------------- Utilidades ----------------
def base_ref(nombre_archivo):
    """Quita la extensión y el sufijo de vista (_D/_LT/_H) -> referencia-color base."""
    stem = os.path.splitext(nombre_archivo)[0].strip()
    up = stem.upper()
    for suf in SUFIJOS_VISTA:
        if up.endswith(suf):
            return stem[: -len(suf)]
    return stem

def orden_vista(nombre_archivo):
    """Ordena: principal (0), detalle (1), lateral-trasera (2), asa (3)."""
    up = os.path.splitext(nombre_archivo)[0].upper()
    if up.endswith("_D"):
        return 1
    if up.endswith("_LT"):
        return 2
    if up.endswith("_H"):
        return 3
    return 0

# Posición "normal" de cada campo (índice de columna en el Excel)
POS = {"nombre": 2, "medidas": 8, "composicion": 9, "precio": 10,
       "color": 13, "web_medidas": 19, "tipo": 20, "asa": 21, "cierre": 22, "estampado": 23}

def extraer(serie):
    """Lee los datos de una fila corrigiendo el descuadre por 'FORRO'.
    Cuando la composición lleva forro, ese texto ocupa una celda extra y desplaza
    +1 todas las columnas a partir de 'Precio venta'."""
    def pos(i):
        try:
            v = serie.iloc[i]
        except Exception:
            return ""
        return "" if pd.isna(v) else str(v).strip()

    precio_cell = pos(POS["precio"])
    shift = "FORRO" in precio_cell.upper()
    off = 1 if shift else 0

    material = pos(POS["composicion"])
    if shift and precio_cell:
        material = (material + " " + precio_cell).strip()     # recupera el forro

    web_medidas = pos(POS["web_medidas"] + off)
    medidas_cruda = pos(POS["medidas"])
    medidas = web_medidas if ("cm" in web_medidas.lower() or " x " in web_medidas.lower()) else medidas_cruda

    return {
        "nombre":    pos(POS["nombre"]),                 # siempre bien (antes del descuadre)
        "medidas":   medidas,
        "material":  material,
        "color":     pos(POS["color"] + off),
        "tipo":      pos(POS["tipo"] + off),
        "asa":       pos(POS["asa"] + off),
        "cierre":    pos(POS["cierre"] + off),
        "estampado": pos(POS["estampado"] + off),
    }

def bloque_datos(d):
    """Construye el bloque de datos reales del producto para pasárselo al modelo."""
    lineas = ["Datos reales del producto (úsalos como verdad, NO inventes nada que no esté aquí o en las fotos):"]
    if d["nombre"]:    lineas.append(f"- Nombre interno: {d['nombre']}")
    if d["tipo"]:      lineas.append(f"- Tipo: {d['tipo']}")
    if d["material"]:  lineas.append(f"- Material/Composición: {d['material']}")
    if d["color"]:     lineas.append(f"- Color: {d['color']}")
    if d["cierre"]:    lineas.append(f"- Cierre: {d['cierre']}")
    if d["asa"]:       lineas.append(f"- Asa: {d['asa']}")
    if d["estampado"]: lineas.append(f"- Estampado: {d['estampado']}")
    if d["medidas"]:   lineas.append(f"- Medidas (solo para orientar el tamaño; NO ponerlas en la descripción): {d['medidas']}")
    return "\n".join(lineas)

def parse_salida(texto):
    """Separa la respuesta del modelo en descripción + 3 titulares."""
    desc = ""
    tits = ["", "", ""]
    m = re.search(r"DESCRIPCI[ÓO]N:\s*(.+?)(?:\n\s*TITULAR\s*1:|\Z)", texto, re.S | re.I)
    if m:
        desc = m.group(1).strip()
    for i in range(3):
        m = re.search(rf"TITULAR\s*{i+1}:\s*(.+?)(?:\n\s*TITULAR\s*{i+2}:|\Z)", texto, re.S | re.I)
        if m:
            tits[i] = m.group(1).strip()
    if not desc and not any(tits):      # por si el modelo no respetó el formato
        desc = texto.strip()
    return desc, tits

def img_block(file_bytes, nombre):
    ext = os.path.splitext(nombre)[1].lower().lstrip(".")
    media = "image/png" if ext == "png" else "image/jpeg"
    return {"type": "image", "source": {"type": "base64", "media_type": media,
            "data": base64.standard_b64encode(file_bytes).decode()}}

def generar(client, prompt, datos_txt, imagenes):
    """Llama a Claude con las fotos + datos + prompt. Devuelve (desc, [titulares])."""
    contenido = [img_block(b, n) for (n, b) in imagenes[:MAX_FOTOS]]
    contenido.append({"type": "text", "text": f"{prompt}\n\n{datos_txt}{INSTRUCCION_FORMATO}"})
    resp = client.messages.create(
        model=MODELO_CLAUDE, max_tokens=1500,
        messages=[{"role": "user", "content": contenido}],
    )
    # Unir solo los bloques de texto (el modelo puede devolver antes bloques de "thinking")
    texto = "".join(getattr(b, "text", "") for b in resp.content if getattr(b, "type", "") == "text")
    return parse_salida(texto)

# ================= INTERFAZ =================
st.set_page_config(page_title="Descripciones de producto · Kbas", page_icon="📝")
st.title("📝 Descripciones de producto")
st.caption("Sube las fotos (se agrupan solas por referencia) y el Excel de datos. Elige el prompt y "
           "la herramienta saca, por cada bolso: descripción + 3 titulares. Luego descargas el Excel.")

def secreto(nombre):
    try:
        return st.secrets.get(nombre, "")
    except Exception:
        return ""

# --- Contraseña (solo se exige si hay 'app_password' en Secrets) ---
_pw = secreto("app_password")
if _pw and not st.session_state.get("_auth_ok"):
    _in = st.text_input("Contraseña", type="password")
    if _in == _pw:
        st.session_state["_auth_ok"] = True
    else:
        if _in:
            st.error("Contraseña incorrecta.")
        st.stop()

# --- Clave de API: SOLO servidor. Si está en Secrets no se muestra nunca. ---
api_key = secreto("anthropic_api_key")
with st.sidebar:
    st.subheader("Motor de IA")
    if api_key:
        st.success("Clave de Claude cargada ✓")
    else:
        api_key = st.text_input("Clave de API de Claude", type="password")
        st.caption("Se consigue en console.anthropic.com → API Keys. Es de pago por uso, aparte de la suscripción.")

c1, c2 = st.columns([1, 1])
with c1:
    eleccion = st.selectbox("Prompt", list(PROMPTS_GUARDADOS.keys()))
with c2:
    st.write(""); st.write("")
    st.caption("Puedes editar el texto de abajo o pegar uno nuevo.")
prompt = st.text_area("Instrucciones para la IA", PROMPTS_GUARDADOS[eleccion], height=200)

fotos = st.file_uploader("Fotos de producto (se agrupan por referencia)", type=["jpg", "jpeg", "png"],
                         accept_multiple_files=True)
excel = st.file_uploader("Excel de datos (opcional, pero recomendado)", type=["xlsx", "xls"])

# Cargar Excel
df = None
if excel is not None:
    try:
        df = pd.read_excel(excel, dtype=str)
        df["_refkey"] = df["Ref."].astype(str).str.strip()
        st.success(f"Excel cargado: {len(df)} filas. Columna de referencia: «Ref.»")
    except Exception as e:
        st.error(f"No pude leer el Excel: {e}")

# Agrupar fotos
grupos = {}
if fotos:
    for f in fotos:
        grupos.setdefault(base_ref(f.name), []).append((f.name, f.getvalue()))
    for k in grupos:
        grupos[k].sort(key=lambda t: orden_vista(t[0]))
    st.info(f"**{len(grupos)} producto(s)** detectado(s) a partir de {len(fotos)} foto(s): "
            + ", ".join(sorted(grupos.keys())))

if grupos:
    if st.button("Generar descripciones", type="primary", disabled=not api_key):
        try:
            import anthropic
            client = anthropic.Anthropic(api_key=api_key)
        except Exception as e:
            st.error(f"No pude iniciar Claude: {e}")
            st.stop()

        filas, barra = [], st.progress(0.0)
        refs = sorted(grupos.keys())
        for i, ref in enumerate(refs):
            d = None
            if df is not None:
                match = df[df["_refkey"] == ref]
                if not match.empty:
                    d = extraer(match.iloc[0])
            datos_txt = bloque_datos(d) if d else "No hay datos de Excel para esta referencia; usa solo las fotos."
            try:
                desc, tits = generar(client, prompt, datos_txt, grupos[ref])
            except Exception as e:
                desc, tits = f"[ERROR: {e}]", ["", "", ""]
            filas.append({
                "Ref.": ref,
                "Tipo": d["tipo"] if d else "",
                "Color": d["color"] if d else "",
                "Medidas": d["medidas"] if d else "",
                "Descripción": desc,
                "Titular 1": tits[0], "Titular 2": tits[1], "Titular 3": tits[2],
                "Nº fotos": len(grupos[ref]),
            })
            barra.progress((i + 1) / len(refs))
        barra.empty()

        res = pd.DataFrame(filas)
        st.success(f"{len(res)} descripción(es) generada(s). Revísalas y descarga el Excel.")
        for _, r in res.iterrows():
            with st.expander(f"{r['Ref.']}  ·  {r['Tipo']}"):
                st.write(r["Descripción"])
                st.write(f"**1.** {r['Titular 1']}")
                st.write(f"**2.** {r['Titular 2']}")
                st.write(f"**3.** {r['Titular 3']}")

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            res.to_excel(w, index=False, sheet_name="Descripciones")
        buf.seek(0)
        st.download_button("⬇ Descargar Excel (referencia → descripción + titulares)", buf,
                           "descripciones.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    elif not api_key:
        st.warning("Introduce tu clave de API de Claude (barra lateral) para poder generar.")
else:
    st.info("Sube las fotos para empezar. El Excel es opcional pero hace las descripciones mucho más precisas.")
