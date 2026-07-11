"""
SACISC - Modulo 04 - Generador de Formato 7 oficial (Excel)
Se ejecuta dentro de un GitHub Action (repository_dispatch).

IMPORTANTE: este script NO reescribe el archivo completo con una libreria
generica (eso borraba el logo y las imagenes). En vez de eso, edita
directamente el XML interno de la celda necesaria, dejando el logo,
los bordes y todo el resto del archivo exactamente igual al original.
"""
import os
import json
import re
import unicodedata
import requests
from xlsx_patch import XlsxSheetPatcher

SUPABASE_URL = os.environ["SUPABASE_URL"]
SUPABASE_SERVICE_KEY = os.environ["SUPABASE_SERVICE_KEY"]
FOLIO_ID = os.environ["FOLIO_ID"]

HEADERS = {
    "apikey": SUPABASE_SERVICE_KEY,
    "Authorization": f"Bearer {SUPABASE_SERVICE_KEY}",
    "Content-Type": "application/json",
}

TEMPLATE_PATH = "plantillas/FORMATO7_template.xlsx"
WORKDIR = "salida"
DIA_ROWS = {"L": 17, "M1": 18, "M2": 19, "J": 20, "V": 21, "S": 22, "D": 23}

CIUDAD_POR_AREA = {
    "SCAD": "AGUA DULCE, VER.",
    "SCEP": "LAS CHOAPAS, VER.",
    "SCCUI": "CUICHAPA, VER.",
    "SCCO": "COATZACOALCOS, VER.",
}

MESES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 5: "MAYO", 6: "JUNIO",
    7: "JULIO", 8: "AGOSTO", 9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
}


def slugificar(texto):
    """Quita acentos y caracteres especiales para que el nombre de archivo
    sea seguro en Supabase Storage (rechaza tildes/enies con error 400)."""
    texto = unicodedata.normalize('NFKD', texto).encode('ascii', 'ignore').decode('ascii')
    texto = re.sub(r'[^A-Za-z0-9_.-]', '_', texto)
    return texto


def fetch_folio(folio_id):
    url = f"{SUPABASE_URL}/rest/v1/pe_folios"
    params = {"id": f"eq.{folio_id}", "select": "*,pe_folio_dias(*)"}
    resp = requests.get(url, headers=HEADERS, params=params)
    resp.raise_for_status()
    data = resp.json()
    if not data:
        raise ValueError(f"No se encontro el folio con id={folio_id}")
    return data[0]


def generar_excel(folio, out_path):
    p = XlsxSheetPatcher(TEMPLATE_PATH)

    p.set_value("AD2", folio["folio"])
    p.set_value("G6", folio["num_semana"])
    if folio.get("fecha_termino_semana"):
        y, m, d = folio["fecha_termino_semana"].split("-")
        p.set_value("N6", int(y))
        p.set_value("P6", int(m))
        p.set_value("R6", int(d))
    p.set_value("F8", "SERVICIOS CORPORATIVOS")
    p.set_value("Q8", 75000)
    p.set_value("Y8", folio["subarea"])
    p.set_value("F10", folio["nombre_trabajador"])
    p.set_value("Y10", folio["ficha"])

    tot_dobletes, tot_comidas = 0, 0
    for dia in folio.get("pe_folio_dias", []):
        r = DIA_ROWS.get(dia["dia_semana"])
        if not r:
            continue
        p.set_value(f"C{r}", dia.get("fecha"))
        salario = folio.get("salario")
        p.set_value(f"D{r}", f"${salario:,.2f}" if salario is not None else None)  # texto formateado con $ y punto, evita coma por config regional
        p.set_value(f"E{r}", folio.get("jornada"))
        p.set_value(f"F{r}", folio.get("nivel"))

        if folio["tipo_pago"] == "DOBLETE":
            if dia.get("horario") == "DESCANSO":
                letras = "DESCANSO"
                cols_descanso = ["K", "L", "M", "N", "O", "P", "Q", "R"]
                for letra, col in zip(letras, cols_descanso):
                    p.set_value(f"{col}{r}", letra)
            elif dia.get("horario"):
                partes = dia["horario"].split("-")
                if len(partes) == 2:
                    p.set_value(f"G{r}", partes[0])
                    p.set_value(f"H{r}", partes[1])
            p.set_value(f"I{r}", dia.get("dobletes") or None)
        elif folio["tipo_pago"] == "TIEMPO_EXTRA":
            if dia.get("horario"):
                partes = dia["horario"].split("-")
                if len(partes) == 2:
                    p.set_value(f"G{r}", partes[0])
                    p.set_value(f"H{r}", partes[1])
            p.set_value(f"K{r}", dia.get("horas") or None)
            p.set_value(f"L{r}", dia.get("minutos") or None)
        elif folio["tipo_pago"] == "INSALUBRE":
            if dia.get("horario"):
                partes = dia["horario"].split("-")
                if len(partes) == 2:
                    p.set_value(f"G{r}", partes[0])
                    p.set_value(f"H{r}", partes[1])
            p.set_value(f"Q{r}", dia.get("horas") or None)
            p.set_value(f"R{r}", dia.get("minutos") or None)

        p.set_value(f"U{r}", dia.get("comidas") or None)
        p.set_value(f"W{r}", "PMXC")  # dato institucional fijo, nunca variable
        p.set_value(f"X{r}", folio.get("partida_presupuestal"))
        p.set_value(f"Y{r}", dia.get("labores_desarrolladas") or "")

        tot_dobletes += dia.get("dobletes") or 0
        tot_comidas += dia.get("comidas") or 0

    p.set_value("I24", tot_dobletes or None)
    p.set_value("U24", tot_comidas or None)

    p.set_value("G26", CIUDAD_POR_AREA.get(folio["area"], ""))
    if folio.get("fecha_reporte"):
        y, m, d = folio["fecha_reporte"].split("-")
        p.set_value("N26", int(d))
        p.set_value("Q26", MESES.get(int(m), ""))
        p.set_value("W26", int(y))

    p.set_value("B30", folio.get("elaboro_nombre"))
    p.set_value("B31", folio.get("elaboro_puesto"))
    p.set_value("K30", folio.get("vobo_nombre"))
    p.set_value("L31", folio.get("vobo_puesto"))
    p.set_value("X30", folio.get("autoriza_nombre"))
    p.set_value("X31", folio.get("autoriza_puesto"))

    p.save(out_path)


def subir_a_storage(local_path, bucket, dest_name):
    with open(local_path, "rb") as f:
        data = f.read()
    url = f"{SUPABASE_URL}/storage/v1/object/{bucket}/{dest_name}"
    headers = dict(HEADERS)
    headers["Content-Type"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    headers["x-upsert"] = "true"
    resp = requests.post(url, headers=headers, data=data)
    resp.raise_for_status()
    return f"{SUPABASE_URL}/storage/v1/object/public/{bucket}/{dest_name}"


def main():
    os.makedirs(WORKDIR, exist_ok=True)
    folio = fetch_folio(FOLIO_ID)

    base = slugificar(f"FORMATO7_{folio['area']}_{folio['subarea']}_{folio['anio']}_{folio['folio']}".replace(" ", "_"))
    xlsx_path = os.path.join(WORKDIR, f"{base}.xlsx")
    generar_excel(folio, xlsx_path)

    bucket = "formato7-generados"
    url_xlsx = subir_a_storage(xlsx_path, bucket, f"{base}.xlsx")

    print(json.dumps({"xlsx": url_xlsx}))


if __name__ == "__main__":
    main()
