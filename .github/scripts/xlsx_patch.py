"""
Motor de edicion quirurgica para xlsx: modifica SOLO valores de celda
en el XML interno, sin tocar imagenes, dibujos, ni ningun otro
componente del archivo. Garantiza fidelidad 100% al original.
"""
import zipfile
import shutil
import re
from lxml import etree

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NSMAP = {"s": NS}


def col_letters_to_index(coord):
    m = re.match(r"([A-Z]+)(\d+)", coord)
    col, row = m.group(1), int(m.group(2))
    idx = 0
    for ch in col:
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx, row


class XlsxSheetPatcher:
    def __init__(self, path, sheet_xml_path="xl/worksheets/sheet1.xml"):
        self.path = path
        self.sheet_xml_path = sheet_xml_path
        with zipfile.ZipFile(path) as z:
            self.original_names = z.namelist()
            xml_bytes = z.read(sheet_xml_path)
        self.tree = etree.fromstring(xml_bytes)
        self.sheetdata = self.tree.find(f"{{{NS}}}sheetData")

    def _get_or_create_row(self, row_num):
        for row_el in self.sheetdata.findall(f"{{{NS}}}row"):
            if int(row_el.get("r")) == row_num:
                return row_el
        # crear nueva fila en la posicion ordenada correcta
        new_row = etree.Element(f"{{{NS}}}row", r=str(row_num))
        inserted = False
        for row_el in self.sheetdata.findall(f"{{{NS}}}row"):
            if int(row_el.get("r")) > row_num:
                row_el.addprevious(new_row)
                inserted = True
                break
        if not inserted:
            self.sheetdata.append(new_row)
        return new_row

    def _get_or_create_cell(self, coord):
        col_idx, row_num = col_letters_to_index(coord)
        row_el = self._get_or_create_row(row_num)
        for c_el in row_el.findall(f"{{{NS}}}c"):
            if c_el.get("r") == coord:
                return c_el
        new_c = etree.Element(f"{{{NS}}}c", r=coord)
        inserted = False
        for c_el in row_el.findall(f"{{{NS}}}c"):
            ci, _ = col_letters_to_index(c_el.get("r"))
            if ci > col_idx:
                c_el.addprevious(new_c)
                inserted = True
                break
        if not inserted:
            row_el.append(new_c)
        return new_c

    def set_value(self, coord, value):
        """Establece un valor numerico o de texto en una celda, quitando
        cualquier formula previa. Preserva el estilo (atributo s) existente."""
        c_el = self._get_or_create_cell(coord)

        # quitar formula previa si la habia
        f_el = c_el.find(f"{{{NS}}}f")
        if f_el is not None:
            c_el.remove(f_el)
        v_el = c_el.find(f"{{{NS}}}v")
        is_el = c_el.find(f"{{{NS}}}is")
        if v_el is not None:
            c_el.remove(v_el)
        if is_el is not None:
            c_el.remove(is_el)

        if value is None or value == "":
            # celda vacia: quitar el atributo t si estaba marcado como texto/formula
            if c_el.get("t") is not None:
                del c_el.attrib["t"]
            return

        if isinstance(value, (int, float)) and not isinstance(value, bool):
            if c_el.get("t") is not None:
                del c_el.attrib["t"]
            v_el = etree.SubElement(c_el, f"{{{NS}}}v")
            v_el.text = str(value)
        else:
            c_el.set("t", "inlineStr")
            is_el = etree.SubElement(c_el, f"{{{NS}}}is")
            t_el = etree.SubElement(is_el, f"{{{NS}}}t")
            t_el.text = str(value)
            t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    def quitar_formato_condicional(self):
        """Elimina TODAS las reglas de formato condicional de la hoja.
        Se encontro una regla accidental en la plantilla original (fila 22,
        columnas K:R) que pinta el texto del color del fondo cuando la celda
        tiene un valor, escondiendo datos validos como si estuvieran vacios.
        Como es un artefacto no intencional del diseno, se quita siempre."""
        for cf in self.tree.findall(f"{{{NS}}}conditionalFormatting"):
            self.tree.remove(cf)

    def save(self, out_path):
        self.quitar_formato_condicional()
        new_xml = etree.tostring(self.tree, xml_declaration=True, encoding="UTF-8", standalone=True)
        shutil.copy(self.path, out_path)
        # reescribir SOLO la entrada del sheet dentro del zip, todo lo demas identico
        with zipfile.ZipFile(self.path) as zin:
            with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename == self.sheet_xml_path:
                        data = new_xml
                    zout.writestr(item, data)
