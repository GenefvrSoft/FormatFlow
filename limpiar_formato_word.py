#!/usr/bin/env python3
import argparse
import sys
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

# Namespaces necesarios para detectar cuadros de texto (no forman parte del
# formato de texto normal, son objetos flotantes/embebidos con su propio XML)
NS_MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS_WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
NS_V = 'urn:schemas-microsoft-com:vml'


def limpiar_run_xml(r_element, mantener_negrita_cursiva=False):
    """
    Elimina todo el formato directo (w:rPr) de un elemento <w:r> trabajando
    directamente sobre el XML. Esto cubre también los runs que están anidados
    dentro de <w:hyperlink>, que paragraph.runs de python-docx NO recorre (es
    una limitación conocida de la librería), y por eso quedaban sin limpiar
    estilos de carácter como "Hipervínculo" (azul y subrayado).
    """
    rPr = r_element.find(qn('w:rPr'))
    if rPr is None:
        return

    if mantener_negrita_cursiva:
        conservar = [el for tag in ('b', 'i') for el in rPr.findall(qn(f'w:{tag}'))]
        r_element.remove(rPr)
        if conservar:
            nueva_rPr = OxmlElement('w:rPr')
            for el in conservar:
                nueva_rPr.append(el)
            r_element.insert(0, nueva_rPr)
    else:
        r_element.remove(rPr)


def limpiar_runs_de_parrafo(paragraph: Paragraph, mantener_negrita_cursiva=False):
    """Limpia TODOS los <w:r> de un párrafo, incluidos los anidados en w:hyperlink."""
    for r_element in paragraph._p.findall('.//' + qn('w:r')):
        limpiar_run_xml(r_element, mantener_negrita_cursiva)


def limpiar_parrafo(paragraph: Paragraph, mantener_negrita_cursiva=False, styles=None):
    """Limpia el formato de un párrafo: estilo, alineación, sangrías, espaciado."""
    limpiar_runs_de_parrafo(paragraph, mantener_negrita_cursiva)

    # Estilo -> Normal (o "List Paragraph" si es una lista, para no perder viñetas/numeración)
    pPr = paragraph._p.find(qn('w:pPr'))
    tiene_numeracion = False
    if pPr is not None:
        numPr = pPr.find(qn('w:numPr'))
        tiene_numeracion = numPr is not None

    if styles is not None:
        try:
            paragraph.style = styles['List Paragraph'] if tiene_numeracion else styles['Normal']
        except KeyError:
            pass  # el documento no tiene ese estilo definido; se deja como está

    paragraph.alignment = None
    paragraph.paragraph_format.left_indent = None
    paragraph.paragraph_format.right_indent = None
    paragraph.paragraph_format.first_line_indent = None
    paragraph.paragraph_format.space_before = None
    paragraph.paragraph_format.space_after = None
    paragraph.paragraph_format.line_spacing = None
    paragraph.paragraph_format.keep_together = None
    paragraph.paragraph_format.keep_with_next = None
    paragraph.paragraph_format.page_break_before = None
    paragraph.paragraph_format.widow_control = None

    # Elimina bordes de párrafo y shading si existen (pBdr, shd)
    if pPr is not None:
        for tag in ('pBdr', 'shd'):
            el = pPr.find(qn(f'w:{tag}'))
            if el is not None:
                pPr.remove(el)
        # Preserva numPr (numeración) pero elimina el resto de pPr "visual"
        for tag in ('spacing', 'ind', 'jc', 'rPr'):
            el = pPr.find(qn(f'w:{tag}'))
            if el is not None:
                pPr.remove(el)


def _contiene_cuadro_texto(elemento):
    """True si el elemento (mc:AlternateContent o w:pict) envuelve un cuadro de texto."""
    return (
        elemento.find(f'.//{{{NS_WPS}}}txbx') is not None
        or elemento.find(f'.//{{{NS_V}}}textbox') is not None
    )


def eliminar_cuadros_texto(xml_root):
    """
    Elimina del árbol XML dado (cuerpo, encabezado, pie o celda de tabla) cualquier
    cuadro de texto: tanto los modernos (mc:AlternateContent con wps:txbx) como los
    de compatibilidad antigua (w:pict con v:textbox). Si el run que lo contenía queda
    vacío, también se elimina ese run.
    Devuelve el número de cuadros de texto eliminados.
    """
    eliminados = 0
    for tag in (f'{{{NS_MC}}}AlternateContent', f'{{{NS_W}}}pict'):
        for elemento in list(xml_root.iter(tag)):
            if not _contiene_cuadro_texto(elemento):
                continue
            padre = elemento.getparent()
            if padre is None:
                continue
            padre.remove(elemento)
            eliminados += 1
            # Si el run (w:r) que contenía el cuadro quedó sin más contenido
            # (aparte de w:rPr), se elimina el run completo para no dejar basura.
            if padre.tag == f'{{{NS_W}}}r':
                hijos_relevantes = [h for h in padre if h.tag != f'{{{NS_W}}}rPr']
                if not hijos_relevantes:
                    abuelo = padre.getparent()
                    if abuelo is not None:
                        abuelo.remove(padre)
    return eliminados


def vaciar_header_footer(parte):
    """
    Elimina por completo el contenido de un encabezado o pie de página (texto,
    imágenes, cuadros de texto, tablas), dejando un único párrafo vacío para
    mantener el documento válido.
    """
    body = parte._element
    parrafos = body.findall(qn('w:p'))
    tablas = body.findall(qn('w:tbl'))

    for tabla in tablas:
        body.remove(tabla)

    # Deja el primer párrafo vacío (sin runs) y elimina el resto
    if parrafos:
        primero = parrafos[0]
        for hijo in list(primero):
            if hijo.tag != qn('w:pPr'):
                primero.remove(hijo)
        for extra in parrafos[1:]:
            body.remove(extra)


def limpiar_bordes_y_sombreado_tabla(tbl_element):
    """
    Elimina bordes y sombreado a nivel de tabla (w:tblBorders, w:shd en tblPr) y a
    nivel de celda (w:tcBorders, w:shd en tcPr), recorriendo el XML directamente
    (tbl_element es el <w:tbl>). Así se cubren también las celdas de continuación
    de una combinación vertical (vMerge), que la abstracción table.rows/cells de
    python-docx no siempre expone una por una. No toca anchos de columna ni
    gridSpan/vMerge, que son estructurales, no visuales.
    """
    tblPr = tbl_element.find(qn('w:tblPr'))
    if tblPr is not None:
        # tblBorders/shd: bordes y sombreado; jc: alineación de la tabla en la página
        for tag in ('tblBorders', 'shd', 'jc'):
            el = tblPr.find(qn(f'w:{tag}'))
            if el is not None:
                tblPr.remove(el)

    for tcPr in tbl_element.findall('.//' + qn('w:tcPr')):
        for tag in ('tcBorders', 'shd'):
            el = tcPr.find(qn(f'w:{tag}'))
            if el is not None:
                tcPr.remove(el)

    # jc en trPr: alineación de la fila en la página (igual de irrelevante visualmente
    # para el texto, pero se elimina para dejar el documento sin ningún rastro de formato)
    for trPr in tbl_element.findall('.//' + qn('w:trPr')):
        el = trPr.find(qn('w:jc'))
        if el is not None:
            trPr.remove(el)


def limpiar_tabla(table, mantener_negrita_cursiva=False, styles=None):
    """Limpia formato de todas las celdas de una tabla (recursivo para tablas anidadas)."""
    limpiar_bordes_y_sombreado_tabla(table._tbl)
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                limpiar_parrafo(paragraph, mantener_negrita_cursiva, styles=styles)
            for tabla_anidada in cell.tables:
                limpiar_tabla(tabla_anidada, mantener_negrita_cursiva, styles=styles)


def limpiar_documento(ruta_entrada, ruta_salida, incluir_headers_footers=False,
                       mantener_negrita_cursiva=False, eliminar_encabezados=False,
                       eliminar_cuadros_texto_doc=False):
    doc = Document(ruta_entrada)

    styles = doc.styles

    if eliminar_cuadros_texto_doc:
        eliminar_cuadros_texto(doc.element.body)

    for paragraph in doc.paragraphs:
        limpiar_parrafo(paragraph, mantener_negrita_cursiva, styles=styles)

    for table in doc.tables:
        limpiar_tabla(table, mantener_negrita_cursiva, styles=styles)

    # Elimina el color de fondo de página (w:background), definido como hermano de
    # w:body en la raíz w:document, NO dentro del cuerpo, por lo que nunca se tocaba.
    fondo = doc.element.find(qn('w:background'))
    if fondo is not None:
        doc.element.remove(fondo)

    for section in doc.sections:
        # IMPORTANTE: cuando el documento tiene activada "primera página diferente"
        # (w:titlePg, expuesto aquí como section.different_first_page_header_footer),
        # existen DOS encabezados y DOS pies de página independientes: el normal
        # (section.header / section.footer) y el de la portada (section.first_page_header
        # / section.first_page_footer). Si no se procesan también estos últimos, todo lo
        # que esté en la portada (logos, cuadros de texto, membretes) sobrevive intacto,
        # que es justo lo que pasaba antes.
        partes = [section.header, section.footer]
        if section.different_first_page_header_footer:
            partes += [section.first_page_header, section.first_page_footer]

        for parte in partes:
            if eliminar_encabezados:
                vaciar_header_footer(parte)
                continue
            if eliminar_cuadros_texto_doc:
                eliminar_cuadros_texto(parte._element)
            if incluir_headers_footers:
                for paragraph in parte.paragraphs:
                    limpiar_parrafo(paragraph, mantener_negrita_cursiva, styles=styles)
                for table in parte.tables:
                    limpiar_tabla(table, mantener_negrita_cursiva, styles=styles)

    doc.save(ruta_salida)


def procesar_carpeta(carpeta_entrada, carpeta_salida, sufijo="",
                      incluir_headers_footers=False, mantener_negrita_cursiva=False,
                      eliminar_encabezados=False, eliminar_cuadros_texto_doc=False):
    carpeta_entrada = Path(carpeta_entrada)
    carpeta_salida = Path(carpeta_salida)
    carpeta_salida.mkdir(parents=True, exist_ok=True)

    archivos = sorted(
        f for f in carpeta_entrada.glob("*.docx")
        if not f.name.startswith("~$")  # ignora archivos temporales de Word abiertos
    )

    if not archivos:
        print(f"No se encontraron archivos .docx en '{carpeta_entrada}'", file=sys.stderr)
        sys.exit(1)

    ok, fallidos = 0, []
    for archivo in archivos:
        nombre_salida = f"{archivo.stem}{sufijo}.docx"
        ruta_salida = carpeta_salida / nombre_salida
        try:
            limpiar_documento(
                archivo,
                ruta_salida,
                incluir_headers_footers=incluir_headers_footers,
                mantener_negrita_cursiva=mantener_negrita_cursiva,
                eliminar_encabezados=eliminar_encabezados,
                eliminar_cuadros_texto_doc=eliminar_cuadros_texto_doc,
            )
            print(f"  ✓ {archivo.name} -> {ruta_salida}")
            ok += 1
        except Exception as e:
            print(f"  ✗ {archivo.name}: {e}", file=sys.stderr)
            fallidos.append(archivo.name)

    print(f"\nListo: {ok}/{len(archivos)} documentos limpiados en '{carpeta_salida}'")
    if fallidos:
        print(f"Fallaron: {', '.join(fallidos)}", file=sys.stderr)


def main():
    parser = argparse.ArgumentParser(
        description="Elimina todo el formato de uno o varios documentos Word (.docx), dejando texto plano."
    )
    parser.add_argument("entrada", help="Archivo .docx, o carpeta con varios .docx, de origen")
    parser.add_argument("salida", help="Archivo .docx de destino, o carpeta de destino si 'entrada' es carpeta")
    parser.add_argument(
        "--incluir-headers-footers",
        action="store_true",
        help="También limpia el formato de encabezados y pies de página",
    )
    parser.add_argument(
        "--mantener-negrita-cursiva",
        action="store_true",
        help="Conserva negrita y cursiva; elimina el resto del formato",
    )
    parser.add_argument(
        "--sufijo",
        default="",
        help="Sufijo a añadir a cada archivo procesado en modo carpeta (ej: _limpio)",
    )
    parser.add_argument(
        "--eliminar-encabezados",
        action="store_true",
        help="Elimina por completo el contenido de encabezados y pies de página (texto, "
             "imágenes, cuadros de texto y tablas), dejándolos en blanco",
    )
    parser.add_argument(
        "--eliminar-cuadros-texto",
        action="store_true",
        help="Elimina los cuadros de texto (text boxes) del documento, incluidos los que "
             "estén dentro de encabezados y pies de página",
    )
    args = parser.parse_args()

    entrada = Path(args.entrada)
    if not entrada.exists():
        print(f"Error: no se encontró '{entrada}'", file=sys.stderr)
        sys.exit(1)

    if entrada.is_dir():
        procesar_carpeta(
            entrada,
            args.salida,
            sufijo=args.sufijo,
            incluir_headers_footers=args.incluir_headers_footers,
            mantener_negrita_cursiva=args.mantener_negrita_cursiva,
            eliminar_encabezados=args.eliminar_encabezados,
            eliminar_cuadros_texto_doc=args.eliminar_cuadros_texto,
        )
    else:
        limpiar_documento(
            entrada,
            args.salida,
            incluir_headers_footers=args.incluir_headers_footers,
            mantener_negrita_cursiva=args.mantener_negrita_cursiva,
            eliminar_encabezados=args.eliminar_encabezados,
            eliminar_cuadros_texto_doc=args.eliminar_cuadros_texto,
        )
        print(f"Documento limpio guardado en: {args.salida}")


if __name__ == "__main__":
    main()
