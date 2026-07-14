"""
SACISC - Generador de Oficio de envio de Formatos 7 a Recursos Humanos.
Se ejecuta dentro de un GitHub Action (repository_dispatch).

Toma un tipo_pago + area + subarea + num_semana + anio, junta todos los
folios que coincidan, y llena la plantilla oficial de oficio (Word) con
la lista de trabajadores -- el diseno/membrete se preserva 100% porque
solo se sustituye texto, nunca se reconstruye el documento desde cero.
"""
import os
import json
import re
import unicodedata
import requests
from oficio_patch import reemplazar_textos, reconstruir_tabla

SUPABASE_URL = os.environ["SUPABASE_URL"]
SUPABASE_SERVICE_KEY = os.environ["SUPABASE_SERVICE_KEY"]
PARAMS = json.loads(os.environ["OFICIO_PARAMS"])

HEADERS = {
    "apikey": SUPABASE_SERVICE_KEY,
    "Authorization": f"Bearer {SUPABASE_SERVICE_KEY}",
    "Content-Type": "application/json",
}

TEMPLATE_PATH = "plantillas/OFICIO_RH_template.docx"
WORKDIR = "salida"

CIUDAD_POR_AREA = {
    "SCAD": "Agua Dulce, Ver.",
    "SCEP": "Las Choapas, Ver.",
    "SCCUI": "Cuichapa, Ver.",
    "SCCO": "Coatzacoalcos, Ver.",
}

MESES = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril", 5: "mayo", 6: "junio",
    7: "julio", 8: "agosto", 9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre",
}

TIPO_PAGO_TEXTO = {
    "DOBLETE": "Dobletes",
    "TIEMPO_EXTRA": "Tiempo extraordinario",
    "INSALUBRE": "Insalubre",
}


def slugificar(texto):
    """Quita acentos y caracteres especiales para que el nombre de archivo
    sea seguro en Supabase Storage (rechaza tildes/enies con error 400)."""
    texto = unicodedata.normalize('NFKD', texto).encode('ascii', 'ignore').decode('ascii')
    texto = re.sub(r'[^A-Za-z0-9_.-]', '_', texto)
    return texto


def fetch_folios():
    url = f"{SUPABASE_URL}/rest/v1/pe_folios"
    params = {
        "area": f"eq.{PARAMS['area']}",
        "subarea": f"eq.{PARAMS['subarea']}",
        "tipo_pago": f"eq.{PARAMS['tipo_pago']}",
        "anio": f"eq.{PARAMS['anio']}",
        "num_semana": f"eq.{PARAMS['num_semana']}",
        "order": "folio.asc",
        "select": "folio,ficha,nombre_trabajador",
    }
    resp = requests.get(url, headers=HEADERS, params=params)
    resp.raise_for_status()
    data = resp.json()
    if not data:
        raise ValueError("No hay folios para esos filtros (area/subarea/tipo_pago/semana).")
    return data


def fetch_firmante(rol):
    url = f"{SUPABASE_URL}/rest/v1/pe_firmantes"
    params = {"rol": f"eq.{rol}", "select": "*", "limit": 1}
    resp = requests.get(url, headers=HEADERS, params=params)
    resp.raise_for_status()
    data = resp.json()
    return data[0] if data else {"nombre": "", "puesto": "", "extension": ""}


def subir_a_storage(local_path, bucket, dest_name):
    with open(local_path, "rb") as f:
        data = f.read()
    url = f"{SUPABASE_URL}/storage/v1/object/{bucket}/{dest_name}"
    headers = dict(HEADERS)
    headers["Content-Type"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    headers["x-upsert"] = "true"
    headers["cache-control"] = "no-cache, no-store, must-revalidate, max-age=0"
    resp = requests.post(url, headers=headers, data=data)
    resp.raise_for_status()
    return f"{SUPABASE_URL}/storage/v1/object/public/{bucket}/{dest_name}"


def main():
    os.makedirs(WORKDIR, exist_ok=True)
    folios = fetch_folios()

    vobo = fetch_firmante("VOBO")
    autoriza = fetch_firmante("AUTORIZA")
    elaboro = fetch_firmante("ELABORO")
    ccp1 = fetch_firmante("CCP1")

    y, m, d = PARAMS["fecha"].split("-")
    fecha_txt = f"{CIUDAD_POR_AREA.get(PARAMS['area'], '')} a  {int(d):02d} de {MESES[int(m)]} de {y}"
    folio_txt = f"DAS-SSAB-URSCS-SSCC-JSCAD-{PARAMS['folio_manual']}-{PARAMS['anio']}"
    tipo_txt = TIPO_PAGO_TEXTO.get(PARAMS["tipo_pago"], PARAMS["tipo_pago"])
    ciudad_area = CIUDAD_POR_AREA.get(PARAMS["area"], "").replace(", Ver.", "")
    asunto_txt = f"Reporte de {tipo_txt.lower()} Semana {PARAMS['num_semana']}, Servicios Corporativos {ciudad_area}-{PARAMS['subarea_corta']}"
    cuerpo_txt = (
        f"Por medio del presente, le solicito gire sus instrucciones a quien corresponda "
        f"para que se realicen los trámites correspondientes para el pago de {tipo_txt} del "
        f"personal de {PARAMS['subarea_corta']} que se enlistan a continuación, correspondientes "
        f"a la semana {PARAMS['num_semana']}."
    )
    elaboro_txt = f"Elaboró: {elaboro.get('nombre','')} Ext. {elaboro.get('extension','')}"

    reemplazos = {
        "Agua Dulce, Ver., a  08 de julio de 2026     ": fecha_txt,
        "DAS-SSAB-URSCS-SSCC-JSCAD-      395       -2026": folio_txt,
        "Mtro. José Antonio Rivera Hernández": PARAMS["destinatario_nombre"],
        "Departamento de Personal Agua Dulce.": PARAMS["destinatario_puesto"],
        "Reporte de tiempo extraordinario Semana 27, Servicios Corporativos Agua Dulce-Áreas Verdes": asunto_txt,
        "Por medio del presente, le solicito gire sus instrucciones a quien corresponda para que se realicen los trámites correspondientes para el pago de Tiempo extraordinario del personal de Áreas Verdes que se enlistan a continuación, correspondientes a la semana 27.": cuerpo_txt,
        "Mtro. Jonatan Eric Reyes Gómez": vobo.get("nombre", ""),
        "Jefe de Servicios Corporativos Agua Dulce": vobo.get("puesto", ""),
        "Ing. José Rogelio Ramírez García": ccp1.get("nombre", ""),
        ". – S.P.A. de la Unidad Regional de Servicios Corporativos Sur.": f". – {ccp1.get('puesto','')}",
        "Lic. Jorge Hernández Landero": autoriza.get("nombre", ""),
        ".- S.P.A. de la Superintendencia de Servicios Corporativos Zona Coatzacoalcos.": f".- {autoriza.get('puesto','')}",
        "Elaboró: Yarumi Leonora Villalobos Kanga Ext. 27-195": elaboro_txt,
    }

    base = slugificar(f"OFICIO_RH_{PARAMS['area']}_{PARAMS['subarea_corta']}_{PARAMS['tipo_pago']}_{PARAMS['anio']}_S{PARAMS['num_semana']}".replace(" ", "_"))
    out_path = os.path.join(WORKDIR, f"{base}.docx")

    reemplazar_textos(TEMPLATE_PATH, out_path, reemplazos)

    filas_tabla = [[f"{i+1:03d}.", f["ficha"], f["nombre_trabajador"]] for i, f in enumerate(folios)]
    reconstruir_tabla(out_path, tabla_idx=1, fila_referencia_idx=1, filas_nuevas=filas_tabla)

    bucket = "oficios-rh-generados"
    url_docx = subir_a_storage(out_path, bucket, f"{base}.docx")

    print(json.dumps({"docx": url_docx}))


if __name__ == "__main__":
    main()
