#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import re
import sys
import unicodedata

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement


# ---------------------------------------------------------------------------
# CONFIGURACIÓN
# ---------------------------------------------------------------------------

# Encabezados de resumen/abstract reconocidos (idioma -> línea de palabras clave)
ABSTRACT_HEADINGS = {
    "RESUMEN": "Palabras clave",
    "ABSTRACT": "Keywords",
    "RESUMO": "Palavras-chave",
    "SUMMARY": "Keywords",
}

# Encabezados de secciones principales del cuerpo que deben llevar el
# prefijo "section:". Se comparan en mayúsculas y sin tildes.
MAIN_SECTIONS = [
    "INTRODUCCION",
    "METODO", "METODOS", "MATERIALES Y METODOS", "METODOLOGIA",
    "RESULTADOS", "RESULTADOS Y DISCUSION",
    "DISCUSION",
    "CONCLUSIONES", "CONCLUSION",
    "ACERCA DE",
]

# Encabezado a partir del cual ya no se procesa nada más (queda intacto)
REFERENCES_HEADINGS = ["REFERENCIAS", "REFERENCES", "BIBLIOGRAFIA"]

# Palabras que delatan una línea de afiliación institucional
AFFIL_KEYWORDS = [
    "universidad", "university", "instituto", "institute", "facultad",
    "college", "école", "ecole", "departamento", "department", "escuela",
    "hospital", "centro", "center", "centre",
]

MESES = {
    "enero": "01", "febrero": "02", "marzo": "03", "abril": "04",
    "mayo": "05", "junio": "06", "julio": "07", "agosto": "08",
    "septiembre": "09", "octubre": "10", "noviembre": "11", "diciembre": "12",
}

EMAIL_RE = re.compile(r"[\w.+-]+@[\w.-]+\.\w+")
ORCID_URL_RE = re.compile(r"https?://orcid\.org/\S+")
ORCID_ID_RE = re.compile(r"\b\d{4}-\d{4}-\d{4}-\d{3}[\dXx]\b")
DATE_LINE_RE = re.compile(
    r"recibid[oa].*?(\d{1,2} de \w+ (?:de )?\d{4}).*?"
    r"acept[oa]d[oa].*?(\d{1,2} de \w+ (?:de )?\d{4}).*?"
    r"public[oa]d[oa].*?(\d{1,2} de \w+ (?:de )?\d{4})",
    re.IGNORECASE | re.DOTALL,
)
SUPERSCRIPT_DIGITS = "⁰¹²³⁴⁵⁶⁷⁸⁹"


# ---------------------------------------------------------------------------
# EDITE AQUÍ SI LA DETECCIÓN AUTOMÁTICA FALLA PARA UN DOCUMENTO EN PARTICULAR
# Deje cada valor en None (o listas vacías) para que se use la detección
# automática; rellene solo lo que necesite forzar manualmente.
# ---------------------------------------------------------------------------
MANUAL_OVERRIDE = {
    "titles": None,          # ej. ["Title[es]:...", "Title[en]:...", "Title[pt]:..."]
    "authors": None,         # ej. [{"name": "...", "orcid": "...", "affil_n": 1}, ...]
    "affiliations": None,    # ej. ["Universidad X. Ciudad, País"]
    "emails": None,          # ej. ["correo1@x.com", "correo2@x.com"]
    "dates": None,           # ej. {"rec": "2026-02-11", "acc": "2026-03-23", "pub": "2026-04-03"}
}


# ---------------------------------------------------------------------------
# UTILIDADES
# ---------------------------------------------------------------------------

def strip_accents(text):
    nfkd = unicodedata.normalize("NFKD", text)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def norm(text):
    """Normaliza un texto para comparaciones robustas (sin tildes, mayúsculas)."""
    return strip_accents(text).strip().upper()


def strip_leading_number(text):
    """Quita un número (normal o superíndice) al inicio de una línea de afiliación."""
    text = text.strip()
    while text and (text[0].isdigit() or text[0] in SUPERSCRIPT_DIGITS):
        text = text[1:]
    return text.strip()


def parse_fecha(txt):
    m = re.search(r"(\d{1,2}) de (\w+) (?:de )?(\d{4})", txt, re.IGNORECASE)
    if not m:
        return None
    d, mes, y = m.group(1), m.group(2).lower(), m.group(3)
    mes_num = MESES.get(mes)
    if not mes_num:
        return None
    return f"{y}-{mes_num}-{int(d):02d}"


# ---------------------------------------------------------------------------
# DETECCIÓN DEL CUERPO (dónde empieza INTRODUCCIÓN, etc.)
# ---------------------------------------------------------------------------

def find_main_section_paragraphs(doc):
    """Devuelve {NOMBRE_NORMALIZADO: objeto Paragraph} de los encabezados
    principales del cuerpo (Introducción, Métodos, Resultados, Discusión,
    Conclusiones...), en el orden en que aparecen."""
    found = {}
    for p in doc.paragraphs:
        t = norm(p.text)
        if not t or len(t) > 60:
            continue
        if t in MAIN_SECTIONS and t not in found:
            found[t] = p
    return found


def find_references_paragraph(doc):
    for p in doc.paragraphs:
        if norm(p.text) in REFERENCES_HEADINGS:
            return p
    return None


# ---------------------------------------------------------------------------
# DETECCIÓN DEL BLOQUE DE ENCABEZADO (título, autores, resúmenes, fechas)
# ---------------------------------------------------------------------------

def find_abstract_blocks(paragraphs):
    """Ubica los bloques RESUMEN/ABSTRACT/RESUMO: encabezado, texto y línea
    de palabras clave. `paragraphs` es la lista completa de Paragraph."""
    blocks = []
    for i, p in enumerate(paragraphs):
        heading = norm(p.text)
        if heading in ABSTRACT_HEADINGS:
            # siguiente párrafo no vacío = texto del resumen
            j = i + 1
            while j < len(paragraphs) and not paragraphs[j].text.strip():
                j += 1
            text = paragraphs[j].text.strip() if j < len(paragraphs) else ""
            # siguiente párrafo no vacío tras el texto = línea de keywords
            k = j + 1
            while k < len(paragraphs) and not paragraphs[k].text.strip():
                k += 1
            kw_line = paragraphs[k].text.strip() if k < len(paragraphs) else ""
            blocks.append({
                "heading": p.text.strip(),
                "heading_idx": i,
                "text": text,
                "keywords_line": kw_line,
                "end_idx": k,
            })
    return blocks


def find_dates(paragraphs, limit_idx):
    for p in paragraphs[:limit_idx]:
        m = DATE_LINE_RE.search(p.text)
        if m:
            return {
                "rec": parse_fecha(m.group(1)),
                "acc": parse_fecha(m.group(2)),
                "pub": parse_fecha(m.group(3)),
            }
    return None


def find_first_person_idx(paragraphs, limit_idx):
    """Primer índice (< limit_idx) cuyo texto contiene un email o un ORCID."""
    for i in range(limit_idx):
        t = paragraphs[i].text
        if EMAIL_RE.search(t) or ORCID_URL_RE.search(t) or ORCID_ID_RE.search(t):
            return i
    return None


def extract_header_block(paragraphs, limit_idx):
    """Extrae títulos, autores, afiliaciones, emails y fechas de los
    párrafos anteriores a `limit_idx` (primer encabezado RESUMEN/ABSTRACT/...).
    """
    dates = find_dates(paragraphs, limit_idx)

    first_person_idx = find_first_person_idx(paragraphs, limit_idx)
    if first_person_idx is None:
        # no se detectaron datos de autor; todo lo anterior son títulos
        titles = [p.text.strip() for p in paragraphs[:limit_idx] if p.text.strip()]
        return {"titles": titles, "authors": [], "affiliations": [], "emails": [], "dates": dates}

    # Absorber hacia atrás las líneas de "nombre" pegadas (sin línea en
    # blanco de por medio) al primer párrafo con email/orcid: eso cubre el
    # formato "nombre / email / orcid" en párrafos separados.
    absorbed_start = first_person_idx
    i = first_person_idx - 1
    while i >= 0 and paragraphs[i].text.strip():
        absorbed_start = i
        i -= 1

    titles = [p.text.strip() for p in paragraphs[:absorbed_start] if p.text.strip()]

    authors = []          # [{"name":..., "orcid":..., "affil_n": int|None}]
    affiliations = []     # [str, ...] en orden de aparición
    emails = []           # [str, ...] en orden de aparición
    current_author = None

    for p in paragraphs[absorbed_start:limit_idx]:
        t = p.text.strip()
        if not t:
            continue
        if DATE_LINE_RE.search(t):
            continue  # ya capturado en `dates`

        email_m = EMAIL_RE.search(t)
        orcid_m = ORCID_URL_RE.search(t) or ORCID_ID_RE.search(t)
        is_affil = any(k in t.lower() for k in AFFIL_KEYWORDS)

        if is_affil and not (email_m or orcid_m):
            affiliations.append(strip_leading_number(t))
            continue

        if orcid_m:
            orcid_val = orcid_m.group(0)
            # texto antes del orcid en la misma línea (formato "combinado")
            name_part = t[:orcid_m.start()].strip()
            if name_part:
                # separar un posible número de afiliación pegado al nombre
                m_num = re.search(r"(\d+|[" + SUPERSCRIPT_DIGITS + r"]+)$", name_part)
                affil_n = None
                if m_num:
                    num_txt = m_num.group(1)
                    affil_n = int(
                        "".join(str(SUPERSCRIPT_DIGITS.index(c)) if c in SUPERSCRIPT_DIGITS else c
                                for c in num_txt)
                    )
                    name_part = name_part[:m_num.start()].strip()
                current_author = {"name": name_part, "orcid": orcid_val, "affil_n": affil_n}
                authors.append(current_author)
            elif current_author is not None:
                current_author["orcid"] = orcid_val
            continue

        if email_m:
            if current_author is not None:
                current_author["email"] = t
            else:
                emails.append(t)
            continue

        # línea "plana": nombre de un nuevo autor
        current_author = {"name": t, "orcid": None, "affil_n": None}
        authors.append(current_author)

    # emails que quedaron adentro de cada autor (formato separado) se
    # trasladan también a la lista `emails`, respetando el orden de autores
    if not emails:
        emails = [a["email"] for a in authors if a.get("email")]

    return {
        "titles": titles,
        "authors": authors,
        "affiliations": affiliations,
        "emails": emails,
        "dates": dates,
    }


# ---------------------------------------------------------------------------
# CONSTRUCCIÓN DEL DOCUMENTO DE SALIDA
# ---------------------------------------------------------------------------

def new_para_before(anchor_elem, parent, style=None):
    p = OxmlElement("w:p")
    anchor_elem.addprevious(p)
    para = docx.text.paragraph.Paragraph(p, parent)
    if style:
        para.style = style
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    return para


def add_run(para, text, bold=False, superscript=False):
    r = para.add_run(text)
    r.bold = bold
    r.font.superscript = superscript
    return r


def flatten_heading(paragraph, new_text, body_text_style):
    for r in list(paragraph.runs):
        r._element.getparent().remove(r._element)
    paragraph.style = body_text_style
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    add_run(paragraph, new_text, bold=False)


def build_ordered_document(src_path, out_path):
    doc = docx.Document(src_path)
    paragraphs = doc.paragraphs
    body_text_style = doc.styles["Body Text"] if "Body Text" in [s.name for s in doc.styles] else doc.styles["Normal"]

    main_sections = find_main_section_paragraphs(doc)
    if not main_sections:
        raise RuntimeError(
            "No se detectó ningún encabezado principal (INTRODUCCIÓN, MÉTODOS, "
            "RESULTADOS, DISCUSIÓN, CONCLUSIONES...). Revise MAIN_SECTIONS."
        )

    # anclaje = primer encabezado principal del cuerpo (normalmente INTRODUCCIÓN)
    anchor_para = min(main_sections.values(), key=lambda p: p._p.getroottree().getpath(p._p))
    # lo anterior puede ser costoso/ambiguo; más simple: usar el índice real
    anchor_idx = min(
        i for i, p in enumerate(paragraphs)
        if norm(p.text) in main_sections
    )
    anchor_para = paragraphs[anchor_idx]

    abstract_blocks = find_abstract_blocks(paragraphs[:anchor_idx])
    header_limit_idx = abstract_blocks[0]["heading_idx"] if abstract_blocks else anchor_idx

    header = MANUAL_OVERRIDE.copy()
    auto = extract_header_block(paragraphs, header_limit_idx)
    for key in ("titles", "authors", "affiliations", "emails", "dates"):
        if header.get(key) is None:
            header[key] = auto[key]

    anchor_elem = anchor_para._p
    parent = anchor_para._parent
    body = anchor_elem.getparent()

    # --- eliminar todo lo anterior al ancla (título, autores, resúmenes...) ---
    to_remove = []
    for child in body.iterchildren():
        if child is anchor_elem:
            break
        to_remove.append(child)
    for child in to_remove:
        body.remove(child)

    # --- reconstruir el bloque de encabezado en el orden del formato ---
    def line(text, bold=False):
        p = new_para_before(anchor_elem, parent, style=body_text_style)
        add_run(p, text, bold=bold)
        new_para_before(anchor_elem, parent, style=body_text_style)  # línea en blanco

    labels = ["es", "en", "pt", "fr", "it"]
    for lbl, title in zip(labels, header["titles"]):
        text = title if title.lower().startswith("title[") else f"Title[{lbl}]:{title}"
        line(text)

    n_affil = len(header["affiliations"])
    for author in header["authors"]:
        p = new_para_before(anchor_elem, parent, style=body_text_style)
        add_run(p, author["name"], bold=True)
        affil_n = author.get("affil_n")
        if affil_n is None:
            affil_n = 1 if n_affil <= 1 else (header["authors"].index(author) + 1)
        add_run(p, str(affil_n), bold=True, superscript=True)
        if author.get("orcid"):
            orcid = author["orcid"]
            if not orcid.startswith("http"):
                orcid = "https://orcid.org/" + orcid
            add_run(p, " " + orcid)
        new_para_before(anchor_elem, parent, style=body_text_style)

    for idx, affil in enumerate(header["affiliations"], start=1):
        p = new_para_before(anchor_elem, parent, style=body_text_style)
        if n_affil > 1 or True:
            add_run(p, str(idx), bold=True, superscript=True)
        add_run(p, affil, bold=True)
        new_para_before(anchor_elem, parent, style=body_text_style)

    for email in header["emails"]:
        line(email)

    for block in abstract_blocks:
        line(block["heading"].upper())
        line(block["text"])
        if block["keywords_line"]:
            line(block["keywords_line"])

    if header["dates"]:
        d = header["dates"]
        if d.get("rec"):
            line(f"date-rec: {d['rec']}")
        if d.get("acc"):
            line(f"date-acc: {d['acc']}")
        if d.get("pub"):
            line(f"date-pub: {d['pub']}")

    # --- normalizar encabezados principales del cuerpo con "section:" ---
    for name, p in main_sections.items():
        flatten_heading(p, f"section:{p.text.strip()}", body_text_style)

    # (REFERENCIAS y todo lo posterior queda intacto)

    doc.save(out_path)
    return header, list(main_sections.keys())


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Uso: python3 ordenar_articulo.py entrada.docx salida.docx")
        sys.exit(1)

    src, out = sys.argv[1], sys.argv[2]
    header, sections = build_ordered_document(src, out)

    print("Documento generado:", out)
    print("Secciones detectadas:", ", ".join(sections))
    print("Títulos:", header["titles"])
    print("Autores:", [(a["name"], a.get("orcid"), a.get("affil_n")) for a in header["authors"]])
    print("Afiliaciones:", header["affiliations"])
    print("Emails:", header["emails"])
    print("Fechas:", header["dates"])
