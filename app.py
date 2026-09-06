"""Brazil Item Coding Studio
Production Version 2.0 - All-in-One
Built for the Brazil team, modeled on India Item Coding Studio V7.1 (Keshav Rajpurohit)

A local desktop automation tool for the Brazil (OGRDS BR) Item Coding workflow,
first configured for LPC 23700 - CHINELOS / SLIPPERS - FLIP-FLOP.

Feature set:
  1. AUTO-TRANSLATE   - the source description (any language) is auto-translated
                         to English and shown next to the original.
  2. GUIDELINE ENGINE  - every characteristic's rule is read directly from the
                         handbook workbook (sheet "OGRDS Characteristic") - PT +
                         EN definition, whether the value is fixed, and examples.
  3. HELP ("?") BUTTON - opens a popup with that field's guideline (PT + EN +
                         examples) PLUS reference photos auto-extracted from the
                         handbook's EXAMPLES sheet, so a coder who doesn't know a
                         char gets a full explanation and a picture in place.
  4. AUTO-SUGGEST      - a per-field Suggest button and a master "Auto-fill from
                         Description" button propose values from the (translated)
                         description following the handbook rules. Suggestions
                         are always shown for review - nothing writes silently.
  5. IDE / reference file linking - load a reference workbook of already-coded
                         items; the tool learns brand/manufacturer pairs and
                         extra dropdown values from it.
  6. Save Changes Online - a concurrency-safe save that re-reads the latest copy
                         of the workbook from disk before writing, so work saved
                         by a teammate on a shared drive is never overwritten.
  7. Local recovery folder - every save also writes a JSON recovery journal next
                         to the allocation file (folder "BrazilItemCodingRecovery")
                         so work is never lost even if the main file is locked,
                         corrupted, or the app crashes. "Recover Local Work"
                         reapplies it.
  8. Ask the Handbook (offline chatbot) - type a free-text question and get an
                         answer built from the handbook's guidelines and scope
                         notes - entirely offline, no internet/API key needed.

Required packages:
    customtkinter>=5.2.2
    openpyxl>=3.1.2
    deep-translator>=1.11.4   (optional - only needed for live translation;
                                the tool still runs without it, translation
                                is simply skipped with a message)

Run:
    pip install customtkinter openpyxl deep-translator
    python Brazil_Item_Coding_Studio_V1.py
"""
from __future__ import annotations
import json, re, warnings, webbrowser, functools, zipfile, io, uuid
from collections import defaultdict
from dataclasses import dataclass, field as dc_field
from pathlib import Path
from datetime import date, datetime
from tkinter import filedialog, messagebox, Menu, Listbox, Toplevel, StringVar, BooleanVar, Text, PhotoImage
import tkinter as tk
from urllib.parse import quote_plus
import customtkinter as ctk
from openpyxl import load_workbook

warnings.filterwarnings("ignore", message="Workbook contains no default style, apply openpyxl's default")

APP = "Brazil Item Coding Studio - All-in-One Version 2.0"

# ---------------------------------------------------------------------------
# Translation backend - THREE tiers, so translation always does *something*
# useful even with no internet connection:
#   1. Live translation via deep-translator (Google Translate) - best quality,
#      handles any language, needs internet.
#   2. Offline domain glossary (PT footwear vocabulary -> EN) - works with
#      zero internet, tailored to this category, used automatically when
#      tier 1 fails or is unavailable.
#   3. Raw passthrough - if all else fails, at least the coder sees the
#      original text instead of an error.
# ---------------------------------------------------------------------------
try:
    from deep_translator import GoogleTranslator
    TRANSLATION_AVAILABLE = True
except Exception:
    TRANSLATION_AVAILABLE = False

_TRANSLATE_CACHE: dict[str, tuple[str, str]] = {}  # original -> (translated, mode)

# Small offline PT -> EN glossary for the flip-flop / footwear domain. Applied
# word-by-word as a fallback when live translation is unavailable, so the
# coder still gets a usable English gist with no internet connection.
OFFLINE_GLOSSARY = {
    "CHINELO":"FLIP-FLOP","CHINELOS":"FLIP-FLOPS","SANDALIA":"SANDAL","SANDALIAS":"SANDALS",
    "SANDÁLIA":"SANDAL","SANDÁLIAS":"SANDALS","FEMININO":"WOMEN'S","FEMININA":"WOMEN'S",
    "MASCULINO":"MEN'S","MASCULINA":"MEN'S","INFANTIL":"CHILDREN'S","ADULTO":"ADULT","ADULTA":"ADULT",
    "UNISSEX":"UNISEX","TAMANHO":"SIZE","TAM":"SIZE","NUMERO":"SIZE","PAR":"PAIR","COR":"COLOR",
    "AZUL":"BLUE","PRETO":"BLACK","PRETA":"BLACK","BRANCO":"WHITE","BRANCA":"WHITE","VERDE":"GREEN",
    "AMARELO":"YELLOW","AMARELA":"YELLOW","VERMELHO":"RED","VERMELHA":"RED","MARROM":"BROWN",
    "ROSA":"PINK","CINZA":"GRAY","DOURADO":"GOLD","DOURADA":"GOLD","PRATEADO":"SILVER","PRATEADA":"SILVER",
    "ESTAMPADO":"PRINTED","ESTAMPADA":"PRINTED","LISO":"SOLID","LISA":"SOLID","LISTRADO":"STRIPED",
    "COMPRIMENTO":"LENGTH","ESPESSURA":"THICKNESS","FORMATO":"SHAPE","CURTA":"SHORT","CURTO":"SHORT",
    "LONGA":"LONG","LONGO":"LONG","FINA":"THIN","FINO":"THIN","GROSSA":"THICK","GROSSO":"THICK",
    "NORMAL":"NORMAL","ESPECIAL":"SPECIAL","COMUM":"COMMON","TRANÇADA":"BRAIDED","TRANÇADO":"BRAIDED",
    "CRUZADA":"CROSSED","CRUZADO":"CROSSED","COM":"WITH","SEM":"WITHOUT","E":"AND","PARA":"FOR",
    "DE":"OF","DA":"OF THE","DO":"OF THE","NOVA":"NEW","NOVO":"NEW","ORIGINAL":"ORIGINAL",
    "DEDO":"TOE","RASTEIRA":"FLAT","RASTEIRINHA":"THONG SANDAL","BEBE":"BABY","BEBÊ":"BABY",
}

def _offline_glossary_translate(value:str)->str:
    words = re.findall(r"[A-Za-zÀ-ÿ0-9/]+|\s+|[^\sA-Za-zÀ-ÿ0-9/]+", value)
    out = []
    for w in words:
        key = w.strip().upper()
        out.append(OFFLINE_GLOSSARY.get(key, w) if key.isalpha() else w)
    return "".join(out)

def translate_to_english(value: str) -> tuple[str, str]:
    """Return (translated_text, mode_label). Never raises. Tries live
    translation first, falls back to the offline glossary, then to the
    original text - so the caller always gets a usable result."""
    value = (value or "").strip()
    if not value:
        return "", ""
    if value in _TRANSLATE_CACHE:
        return _TRANSLATE_CACHE[value]
    if TRANSLATION_AVAILABLE:
        for attempt in range(2):  # one retry - transient network hiccups are common
            try:
                translator = GoogleTranslator(source="auto", target="en")
                translated = translator.translate(value)
                if translated and translated.strip():
                    result = (translated.strip(), "auto-translated")
                    _TRANSLATE_CACHE[value] = result
                    return result
            except Exception:
                continue
    # Live translation unavailable/failed - fall back to the offline glossary.
    glossary_result = _offline_glossary_translate(value)
    if upper(glossary_result) != upper(value):
        result = (glossary_result, "offline glossary (no internet / translator library missing)")
    else:
        result = (value, "translation unavailable - showing original text")
    _TRANSLATE_CACHE[value] = result
    return result

# ---------------------------------------------------------------------------
# Small text helpers (same conventions as the India tool)
# ---------------------------------------------------------------------------
def text(v): return "" if v is None else str(v).strip()
def upper(v): return re.sub(r"\s+", " ", text(v)).upper().strip()

PROPOSED_PREFIX = "Proposed Value - "
ID_COLUMNS = {"lpc_id", "lpc_dscr", "External Code", "Product Title"}
WORKFLOW_COLUMNS = ["Allocation", "Status", "Date", "Link", "Comment"]
DESCRIPTION_HEADER = "Product Title"
ITEM_DESC_CHAR = "#BR LOC 1000001 : ITM_DESC"

@dataclass
class Field:
    name: str
    source_col: int | None
    proposed_col: int

@dataclass
class Guideline:
    name: str
    pt_def: str = ""
    en_def: str = ""
    fixed: str = ""
    pt_examples: str = ""
    en_examples: str = ""
    extra_notes: str = ""     # extended explanation mined from the EXAMPLES sheet
    enum_values: tuple = ()   # the enumerated valid-value list, when the guideline lists one
    meta: dict = dc_field(default_factory=dict)  # Group/SubGroup/Category/Country, from the master char list (Sheet3)

# ---------------------------------------------------------------------------
# Handbook loader - generic parser for the "OGRDS Characteristic" sheet, the
# "Definition" scope sheet, the "EXAMPLES" worked-examples sheet (including
# its embedded reference photos), and the master characteristic list (the
# sheet with Code/Description/Group/SubGroup/Category/Country columns).
# Works for any Brazil handbook workbook that follows the same layout:
#   Characteristic | Como preencher/Definicao | How to populate/Definition |
#   Value Fixed? | Exemplos de Valores | Values examples
# ---------------------------------------------------------------------------
@dataclass
class HandbookData:
    guidelines: dict = dc_field(default_factory=dict)   # name -> Guideline
    scope_notes: list = dc_field(default_factory=list)  # free-text scope/definition/example paragraphs
    images: dict = dc_field(default_factory=dict)       # (LABEL_KEY) -> list[bytes]
    metadata: dict = dc_field(default_factory=dict)     # name -> {code, group, subgroup, category, country}

def _sheet_guidelines(ws) -> dict[str, Guideline]:
    guidelines: dict[str, Guideline] = {}
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return guidelines
    header = [upper(c) for c in rows[0]]
    def col(*names):
        for n in names:
            if n in header:
                return header.index(n)
        return None
    c_name = col("CHARACTERISTIC")
    c_pt_def = col("COMO PREENCHER/DEFINIÇÃO", "COMO PREENCHER/DEFINICAO")
    c_en_def = col("HOW TO POPULATE/DEFINITION")
    c_fixed = col("VALUE FIXED?")
    c_pt_ex = col("EXEMPLOS DE VALORES", "EXEMPLOS DE VALORES ")
    c_en_ex = col("VALUES EXAMPLES")
    for row in rows[1:]:
        if c_name is None or c_name >= len(row):
            continue
        name = text(row[c_name])
        if not name:
            continue
        def cell(idx):
            return text(row[idx]) if idx is not None and idx < len(row) else ""
        en_ex = cell(c_en_ex)
        enum_values = ()
        # When the "Values examples" cell lists several short newline-separated
        # tokens, treat it as the field's full enumerated value set (e.g. VERSAO's
        # S/V, PERSON, TOP, BRASIL...) so it can seed the dropdown autocomplete.
        lines = [upper(l) for l in re.split(r"[\n;]", en_ex) if upper(l)]
        if 1 < len(lines) <= 15 and all(len(l) <= 40 for l in lines):
            enum_values = tuple(dict.fromkeys(lines))
        guidelines[name] = Guideline(
            name=name, pt_def=cell(c_pt_def), en_def=cell(c_en_def),
            fixed=cell(c_fixed), pt_examples=cell(c_pt_ex), en_examples=en_ex,
            enum_values=enum_values,
        )
    return guidelines

def _sheet_scope_notes(ws) -> list[str]:
    notes = []
    for row in ws.iter_rows(values_only=True):
        for cell in row:
            value = text(cell)
            if value and len(value) > 15:
                notes.append(value)
    return notes

def _sheet_examples_notes(ws, guidelines: dict[str, Guideline]) -> list[str]:
    """Mines the EXAMPLES sheet for (a) every worked-example paragraph, used to
    enrich the offline chatbot's corpus, and (b) longer descriptive text found
    near a '#BR LOC ... : CHAR VALUE' label cell, which gets appended onto that
    characteristic's guideline as extended, example-grounded explanation."""
    cells = []
    for row in ws.iter_rows():
        for c in row:
            v = text(c.value)
            if v:
                cells.append((c.row, c.column, v))
    notes = [v for (_, _, v) in cells if len(v) > 25]
    for lrow, lcol, ltext in cells:
        if not ltext.upper().startswith("#BR LOC"):
            continue
        base = next((g for g in guidelines if ltext.upper().startswith(upper(g))), None)
        if not base:
            continue
        candidates = [(r2, c2, t) for (r2, c2, t) in cells
                      if lrow < r2 <= lrow + 15 and abs(c2 - lcol) <= 3 and len(t) > 40]
        if not candidates:
            continue
        r2, c2, best = min(candidates, key=lambda item: (item[0] - lrow) * 10 + abs(item[1] - lcol))
        tag = ltext.split(":")[-1].strip() or ltext
        addition = f"[{tag}] {best}"
        g = guidelines[base]
        if addition not in g.extra_notes:
            g.extra_notes = (g.extra_notes + "\n" if g.extra_notes else "") + addition
    return notes

def _sheet3_metadata(ws) -> dict[str, dict]:
    """Parses the master characteristic list (Code / Description / Group /
    SubGroup / Category / Country) into name -> metadata, giving each field
    extra technical context (e.g. whether it's Brazil-only or multi-country)."""
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return {}
    header = [upper(c) for c in rows[0]]
    idx = {h: i for i, h in enumerate(header)}
    if "DESCRIPTION" not in idx:
        return {}
    out = {}
    for r in rows[1:]:
        di = idx["DESCRIPTION"]
        name = text(r[di]) if di < len(r) else ""
        if not name:
            continue
        def cell(key):
            i = idx.get(key)
            return text(r[i]) if i is not None and i < len(r) else ""
        out[name] = {"code": cell("CODE"), "group": cell("GROUP"), "subgroup": cell("SUBGROUP"),
                     "category": cell("CATEGORY"), "country": cell("COUNTRY")}
    return out

def _extract_handbook_images(path) -> dict[str, list[bytes]]:
    """Best-effort extraction of the handbook's embedded example photos, mapped
    to the nearest '#BR LOC ... : CHAR VALUE' label above/left of each picture
    (this matches how the handbook's EXAMPLES sheet is laid out: a label cell
    followed a few rows below by one or more photos in the same column band)."""
    images: dict[str, list[bytes]] = defaultdict(list)
    try:
        with zipfile.ZipFile(path) as zf:
            names = set(zf.namelist())
            if "xl/workbook.xml" not in names:
                return images
            workbook_xml = zf.read("xl/workbook.xml").decode("utf-8", "ignore")
            rels_xml = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8", "ignore")
            sheet_to_rid = dict(re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', workbook_xml))
            rid_to_target = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels_xml))
            target = None
            for sheet_name, rid in sheet_to_rid.items():
                if sheet_name.upper() == "EXAMPLES":
                    target = rid_to_target.get(rid)
                    break
            if not target:
                return images
            sheet_path = "xl/" + target.replace("../", "")
            sheet_file = target.split("/")[-1]
            sheet_rels_path = f"xl/worksheets/_rels/{sheet_file}.rels"
            if sheet_rels_path not in names:
                return images
            sheet_rels = zf.read(sheet_rels_path).decode("utf-8", "ignore")
            m = re.search(r'Target="([^"]*drawing[^"]+)"', sheet_rels)
            if not m:
                return images
            drawing_path = "xl/" + m.group(1).replace("../", "")
            drawing_file = drawing_path.split("/")[-1]
            drawing_rels_path = f"xl/drawings/_rels/{drawing_file}.rels"
            drawing_xml = zf.read(drawing_path).decode("utf-8", "ignore")
            drawing_rels = zf.read(drawing_rels_path).decode("utf-8", "ignore") if drawing_rels_path in names else ""
            rid_to_media = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', drawing_rels))
            anchors = []
            for a in re.findall(r"<xdr:twoCellAnchor.*?</xdr:twoCellAnchor>", drawing_xml, re.S):
                fr = re.search(r"<xdr:from><xdr:col>(\d+)</xdr:col>.*?<xdr:row>(\d+)</xdr:row>", a, re.S)
                rid = re.search(r'r:embed="(rId\d+)"', a)
                if fr and rid:
                    anchors.append((int(fr.group(2)) + 1, int(fr.group(1)) + 1, rid.group(1)))  # 1-indexed row, col
            # Load EXAMPLES cell text via openpyxl for label matching.
            wb = load_workbook(path, read_only=True, data_only=True)
            examples_ws = wb["EXAMPLES"] if "EXAMPLES" in wb.sheetnames else None
            labels = {}
            if examples_ws is not None:
                for row in examples_ws.iter_rows():
                    for cell in row:
                        v = text(cell.value)
                        if v and v.upper().startswith("#BR LOC"):
                            labels[(cell.row, cell.column)] = v
            wb.close()
            for img_row, img_col, rid in anchors:
                media = rid_to_media.get(rid)
                if not media:
                    continue
                media_path = "xl/" + media.replace("../", "")
                if media_path not in names:
                    continue
                best_label, best_dist = None, 999
                for (lrow, lcol), label_text in labels.items():
                    if lrow < img_row <= lrow + 15 and abs(lcol - img_col) <= 2:
                        dist = (img_row - lrow) * 10 + abs(lcol - img_col)
                        if dist < best_dist:
                            best_dist, best_label = dist, label_text
                if best_label:
                    key = upper(re.sub(r"\s+", " ", best_label))
                    images[key].append(zf.read(media_path))
    except Exception:
        pass  # image extraction is a convenience feature - never block the app
    return images

def load_handbook(path) -> HandbookData:
    wb = load_workbook(path, read_only=True, data_only=True)
    sheet_name = next((s for s in wb.sheetnames if "OGRDS CHARACTERISTIC" in s.upper()), None)
    if sheet_name is None:
        wb.close()
        raise ValueError("Sheet 'OGRDS Characteristic' was not found in the handbook workbook.")
    guidelines = _sheet_guidelines(wb[sheet_name])
    definition_sheet = next((s for s in wb.sheetnames if "DEFINITION" in s.upper()), None)
    scope_notes = _sheet_scope_notes(wb[definition_sheet]) if definition_sheet else []
    examples_sheet = next((s for s in wb.sheetnames if "EXAMPLES" in s.upper()), None)
    if examples_sheet:
        scope_notes += _sheet_examples_notes(wb[examples_sheet], guidelines)
    # The master characteristic list isn't reliably named - detect it by its
    # header row (Code/Description/Group/SubGroup/Category/Country) instead.
    metadata = {}
    for sname in wb.sheetnames:
        if sname in {sheet_name, definition_sheet, examples_sheet}:
            continue
        try:
            first_row = [upper(c) for c in next(wb[sname].iter_rows(max_row=1, values_only=True), [])]
        except StopIteration:
            continue
        if {"CODE", "DESCRIPTION", "GROUP"}.issubset(set(first_row)):
            metadata = _sheet3_metadata(wb[sname])
            break
    wb.close()
    for name, meta in metadata.items():
        if name in guidelines:
            guidelines[name].meta = meta
    images = _extract_handbook_images(path)
    return HandbookData(guidelines=guidelines, scope_notes=scope_notes, images=images, metadata=metadata)

# ---------------------------------------------------------------------------
# SearchEntry - fast searchable dropdown (unchanged from the India tool; this
# widget is generic and not specific to any country's field list).
# ---------------------------------------------------------------------------
class SearchEntry(ctk.CTkEntry):
    """Fast searchable dropdown with reliable one-click opening."""
    def __init__(self,master,values=None,strict=False,on_pick=None,on_next=None,on_manual_open=None,**kwargs):
        self.var=kwargs.pop("textvariable",StringVar())
        super().__init__(master,textvariable=self.var,**kwargs)
        self.values=[];self.strict=strict;self.on_pick=on_pick;self.on_next=on_next;self.on_manual_open=on_manual_open
        self.popup=None;self.listbox=None;self.visible_values=[];self.active_index=0;self.show_all_on_open=False;self.keyboard_choice=False
        self.last_valid="";self._filter_job=None;self._auto_job=None;self._close_job=None
        self.set_values(values or [])
        self.bind("<Button-1>",self.manual_open,add="+")
        self.bind("<FocusIn>",self.focus_open,add="+")
        self.bind("<KeyRelease>",self.key_release,add="+")
        self.bind("<Return>",self.accept_next);self.bind("<Tab>",self.accept_next)
        self.bind("<Down>",lambda e:self.move_active(1));self.bind("<Up>",lambda e:self.move_active(-1))
        self.bind("<Escape>",lambda e:self.close())
        try:
            self._entry.bind("<Button-1>",self.manual_open,add="+")
            self._entry.bind("<FocusIn>",self.focus_open,add="+")
            self._entry.bind("<KeyRelease>",self.key_release,add="+")
            self._entry.bind("<Return>",self.accept_next)
            self._entry.bind("<Tab>",self.accept_next)
            self._entry.bind("<Down>",lambda e:self.move_active(1))
            self._entry.bind("<Up>",lambda e:self.move_active(-1))
            self._entry.bind("<Escape>",lambda e:self.close())
        except Exception:pass
        self.bind("<FocusOut>",self.commit_on_focus_out,add="+")
        self.bind("<FocusOut>",self.delayed_close,add="+")
    def set_values(self,values):
        normalized_values=sorted({upper(x) for x in values if upper(x)})
        if normalized_values==self.values:return
        self.values=normalized_values
        if self.popup:self.schedule_refresh(0)
    def matches(self):
        if self.show_all_on_open:return self.values
        q=upper(self.var.get())
        if not q:return self.values
        starts=[];contains=[]
        for value in self.values:
            if value.startswith(q):starts.append(value)
            elif q in value:contains.append(value)
        return starts+contains
    def cancel_close(self):
        if self._close_job:
            try:self.after_cancel(self._close_job)
            except:pass
        self._close_job=None
    def manual_open(self,event=None):
        self.cancel_close();self.show_all_on_open=True
        if self.on_manual_open:self.on_manual_open()
        self.after_idle(lambda:self.open(force=True))
    def focus_open(self,event=None):
        self.cancel_close()
        if not upper(self.var.get()):self.after_idle(lambda:self.open(force=False))
    def open(self,force=False):
        if str(self.cget("state"))=="disabled" or not self.values:return
        if not force and upper(self.var.get()):return
        self.cancel_close()
        if self.popup and self.popup.winfo_exists():
            self.schedule_refresh(0);return
        self.popup=Toplevel(self);self.popup.overrideredirect(True);self.popup.attributes("-topmost",True)
        self.popup.configure(bg="#FFFFFF")
        self.listbox=Listbox(self.popup,font=("Segoe UI",11),activestyle="none",selectmode="browse",
                             selectbackground="#2F80C1",selectforeground="#FFFFFF",
                             bg="#FFFFFF",fg="#153B53",bd=1,relief="solid",highlightthickness=0,
                             exportselection=False)
        self.listbox.pack(fill="both",expand=True)
        self.listbox.bind("<ButtonPress-1>",self.mouse_select)
        self.listbox.bind("<ButtonRelease-1>",lambda e:"break")
        self.listbox.bind("<MouseWheel>",self.mouse_wheel)
        self.listbox.bind("<Return>",self.accept_next);self.listbox.bind("<Tab>",self.accept_next)
        self.listbox.bind("<Down>",lambda e:self.move_active(1));self.listbox.bind("<Up>",lambda e:self.move_active(-1))
        self.active_index=0;self.keyboard_choice=False;self.refresh()
        try:self._entry.focus_set()
        except:self.focus_set()
    def schedule_refresh(self,delay=35):
        if self._filter_job:
            try:self.after_cancel(self._filter_job)
            except:pass
        self._filter_job=self.after(delay,self.refresh)
    def refresh(self):
        self._filter_job=None
        if not self.popup or not self.listbox:return
        self.visible_values=self.matches();self.active_index=max(0,min(self.active_index,len(self.visible_values)-1))
        self.listbox.delete(0,"end")
        if self.visible_values:self.listbox.insert("end",*self.visible_values)
        self.paint_active();self.place_below()
    def place_below(self):
        if not self.popup:return
        self.update_idletasks();app=self.winfo_toplevel()
        width=self.winfo_width();x=self.winfo_rootx();y=self.winfo_rooty()+self.winfo_height()+2
        app_bottom=app.winfo_rooty()+app.winfo_height()-74;available=app_bottom-y
        if available<95:
            try:
                owner=self.master
                while owner is not None and not hasattr(owner,"_parent_canvas"):owner=getattr(owner,"master",None)
                if owner is not None:owner._parent_canvas.yview_scroll(4,"units")
                app.update_idletasks();x=self.winfo_rootx();y=self.winfo_rooty()+self.winfo_height()+2;available=app_bottom-y
            except Exception:pass
        rows=max(1,min(5,len(self.visible_values)));height=max(35,min(rows*27+4,max(35,available)))
        self.popup.geometry(f"{width}x{height}+{x}+{y}");self.popup.lift()
    def paint_active(self):
        if not self.listbox or not self.visible_values:return
        self.listbox.selection_clear(0,"end");self.listbox.selection_set(self.active_index)
        self.listbox.activate(self.active_index);self.listbox.see(self.active_index)
    def move_active(self,step):
        self.cancel_close()
        if not self.popup:self.open(force=True)
        if not self.visible_values:return "break"
        self.active_index=(self.active_index+step)%len(self.visible_values);self.keyboard_choice=True;self.paint_active();return "break"
    def key_release(self,event):
        if event.keysym in {"Return","Tab","Escape","Down","Up","Shift_L","Shift_R"}:return
        self.show_all_on_open=False;self.keyboard_choice=False;self.cancel_close();self.active_index=0;self.open(force=True);self.schedule_refresh()
    def mouse_wheel(self,event):
        if not self.visible_values:return "break"
        return self.move_active(-1 if event.delta>0 else 1)
    def mouse_select(self,event=None):
        self.cancel_close()
        if self.listbox and self.visible_values:
            index=self.listbox.nearest(event.y) if event is not None else (self.listbox.curselection()[0] if self.listbox.curselection() else 0)
            if 0<=index<len(self.visible_values):
                self.active_index=index;self.keyboard_choice=False;self.paint_active()
                self.after_idle(lambda v=self.visible_values[index]:self.pick(v))
        return "break"
    def pick(self,value):
        value=upper(value)
        if self.strict and value not in self.values:return
        self.var.set(value);self.last_valid=value;self.close()
        if self.on_pick:self.after_idle(lambda v=value:self.on_pick(v))
    def accept_next(self,event=None):
        typed=upper(self.var.get())
        highlighted=(self.visible_values[self.active_index]
                     if self.popup and self.visible_values and 0<=self.active_index<len(self.visible_values) else "")
        value=highlighted if self.keyboard_choice and highlighted else typed
        if self.strict and value not in self.values:self.var.set(self.last_valid)
        else:
            self.var.set(value);self.last_valid=value
            if self.on_pick:self.after_idle(lambda v=value:self.on_pick(v))
        self.keyboard_choice=False;self.close();self.move_next();return "break"
    def move_next(self):
        if self.on_next:self.after(5,self.on_next)
    def commit_on_focus_out(self,event=None):
        value=upper(self.var.get())
        if value and value!=self.last_valid:
            if self.strict and value not in self.values:
                self.var.set(self.last_valid)
            else:
                self.last_valid=value
                if self.on_pick:self.after_idle(lambda v=value:self.on_pick(v))
    def delayed_close(self,event=None):
        self.cancel_close();self._close_job=self.after(250,self.close_if_outside)
    def close_if_outside(self):
        self._close_job=None
        if not self.popup:return
        focus=self.focus_get();internal=getattr(self,"_entry",None)
        if focus in {self,internal,self.listbox}:return
        try:
            x,y=self.winfo_pointerx(),self.winfo_pointery();px,py=self.popup.winfo_rootx(),self.popup.winfo_rooty()
            if px<=x<=px+self.popup.winfo_width() and py<=y<=py+self.popup.winfo_height():return
        except Exception:pass
        self.close()
    def close(self):
        self.cancel_close()
        if self._filter_job:
            try:self.after_cancel(self._filter_job)
            except:pass
        self._filter_job=None
        if self.popup:
            try:self.popup.destroy()
            except:pass
        self.popup=None;self.listbox=None;self.visible_values=[];self.active_index=0

# ---------------------------------------------------------------------------
# ScrollableFrame - a hand-built replacement for CTkScrollableFrame.
#
# CTkScrollableFrame's own global mouse-wheel binding is unreliable with
# laptop trackpad two-finger scroll on several CustomTkinter versions (the
# reported bug). This class uses the classic, dependable Tk pattern instead:
# a raw Canvas + inner Frame, with the wheel binding added ONLY while the
# pointer is actually over this canvas (via <Enter>/<Leave>) and removed the
# moment it leaves. Because the binding only exists while hovering, there is
# never any ambiguity about which scrollable region should move, so it works
# identically for a mouse wheel or a trackpad two-finger swipe, on Windows,
# macOS and Linux.
# ---------------------------------------------------------------------------
class ScrollableFrame(ctk.CTkFrame):
    """A hand-built replacement for CTkScrollableFrame.

    IMPORTANT DESIGN NOTE: an earlier version of this class tried to bind the
    mouse wheel only while the pointer hovered the canvas/content widgets
    directly (via <Enter>/<Leave>). That approach is broken whenever the
    content is fully tiled with child widgets (rows, labels, entries) - which
    it always is here - because <Enter>/<Leave> only fire for whichever
    widget is literally topmost under the cursor, so the canvas/content
    itself almost never receives them and the wheel binding never activates.
    That is exactly why "click-and-drag the scrollbar works but the wheel/
    trackpad doesn't" was reported.

    The fix: bind the wheel event ONCE, globally, at the class level (not per
    instance), and on every wheel event hit-test the real cursor position
    with winfo_containing() to find whichever widget is actually under it,
    then walk UP that widget's parent chain to see whether it lives inside
    any registered ScrollableFrame's canvas. This works no matter how deeply
    nested the widget under the cursor is - buttons, entries, labels, rows -
    because the parent-chain walk always finds the owning canvas.
    """
    _instances: list = []
    _bound = False

    def __init__(self,master,fg_color="transparent",**kwargs):
        super().__init__(master,fg_color=fg_color,**kwargs)
        bg=self._apply_appearance_mode(ctk.ThemeManager.theme["CTkFrame"]["fg_color"]) if fg_color=="transparent" else fg_color
        self.canvas=tk.Canvas(self,highlightthickness=0,bg=bg,bd=0)
        self.vsb=ctk.CTkScrollbar(self,orientation="vertical",command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.canvas.pack(side="left",fill="both",expand=True)
        self.vsb.pack(side="right",fill="y")
        self.content=ctk.CTkFrame(self.canvas,fg_color=fg_color)
        self.window_id=self.canvas.create_window((0,0),window=self.content,anchor="nw")
        self.content.bind("<Configure>",lambda e:self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>",self._on_canvas_resize)
        ScrollableFrame._instances.append(self)
        self.bind("<Destroy>",self._on_destroy,add="+")
        if not ScrollableFrame._bound:
            top=self.winfo_toplevel()
            top.bind_all("<MouseWheel>",ScrollableFrame._dispatch,add="+")  # Windows + macOS (incl. trackpad)
            top.bind_all("<Button-4>",ScrollableFrame._dispatch,add="+")    # Linux scroll up
            top.bind_all("<Button-5>",ScrollableFrame._dispatch,add="+")    # Linux scroll down
            ScrollableFrame._bound=True
    def _on_destroy(self,event):
        try:ScrollableFrame._instances.remove(self)
        except ValueError:pass
    def _on_canvas_resize(self,event):
        self.canvas.itemconfig(self.window_id,width=event.width)
    @classmethod
    def _dispatch(cls,event):
        try:
            under=event.widget.winfo_containing(event.x_root,event.y_root)
        except Exception:
            under=None
        if under is None:
            return None
        for inst in list(cls._instances):
            node=under
            depth=0
            while node is not None and depth<60:
                if node is inst.canvas or node is inst.content:
                    inst._scroll(event)
                    return "break"
                node=getattr(node,"master",None); depth+=1
        return None  # cursor isn't over any managed scroll area - let default handling proceed
    def _scroll(self,event):
        if getattr(event,"num",None)==4:steps=-3
        elif getattr(event,"num",None)==5:steps=3
        else:
            delta=getattr(event,"delta",0)
            steps=-3 if delta>0 else 3
        self.canvas.yview_scroll(int(steps),"units")

# ---------------------------------------------------------------------------
# Rule / auto-suggest engine for LPC 23700 - CHINELOS (Slippers/Flip-Flops).
# Every rule below is taken directly from:
#   - 23700_CHINELOS_HANDBOOK_CORRETO.xlsx ("OGRDS Characteristic" sheet)
#   - BRC-CLT.pptx (23700 SLIPPERS - FLIP-FLOP guideline deck)
#   - Learning_Priority_2 (LATAM BR).docx (coding-quality feedback rules)
#
# Design: each field has a suggest_<field> function that receives a Context
# object (the description, its English translation, tokens, and whatever has
# already been proposed for other fields this session) and returns either a
# suggested value (string) or None if it cannot confidently guess. Nothing
# here ever writes to the workbook directly - App.suggest_field() takes the
# return value and puts it in the on-screen box for the coder to review.
# ---------------------------------------------------------------------------

BRAND_KEY = "#BR LOC 90090 : MARCA"
BRAND_G2G_KEY = "#BR LOC 80090 : MARCA G2G"
BRAND1_KEY = "BRAND 1"
SUBBRAND_KEY = "#BR LOC 35 : SUBMARCA"
MANUFACTURER_KEY = "#BR LOC 90080 : FABRICANTE"
MANUFACTURER_G2G_KEY = "#BR LOC 80080 : FABRICANTE G2G"
BOI_KEY = "BRAND OWNER - INTERNATIONAL"
MODULE_KEY = "MODULE"
ITM_CLASE_KEY = "#BR LOC 1000009: ITM_CLASE_PRODUCTO"
FORMATO_KEY = "#BR LOC 1593 : FORMATO"
FORMATO_TRANSF_KEY = "#BR LOC 4 : FORMATO DE TRANSFERENCIA"
ITM_CONTENIDO_KEY = "#BR LOC 1000003 : ITM_CONTENIDO"
TAMANHO_KEY = "#BR LOC 36 : TAMANHO"
COMPRIMENTO_KEY = "#BR LOC 1596 : COMPRIMENTO"
ESPESSURA_KEY = "#BR LOC 1600 : ESPESSURA"
COLORACAO_KEY = "#BR LOC 1603 : COLORAÇÃO"
VERSAO_KEY = "#BR LOC 1671 : VERSAO"
EMBALAGEM_KEY = "#BR LOC 300001 : EMBALAGEM"
FT_CONV_KEY = "#BR LOC 301010 : FT_CONV UNIDADES"
INTERNA_MT_KEY = "#BR LOC 2025 : INTERNA CAMPO MT"
PROMO_TYPE_KEY = "PROMOTIONAL ACTIVITY - TYPE"
PROMO_OFFER_KEY = "PROMOTIONAL OFFER"
GTI_PACK_KEY = "GLOBAL TOTAL ITEMS IN PACK"
GN_EACH_BASE_KEY = "GLOBAL NUMBER IN EACH PACK - BASE"
GN_EACH_ACTUAL_KEY = "GLOBAL NUMBER IN EACH PACK - ACTUAL"
GTP_MULTI_KEY = "GLOBAL TOTAL PACKS IN MULTIPACK"
GN_MULTI_BASE_KEY = "GLOBAL NUMBER IN MULTIPACK - BASE"
GN_MULTI_ACTUAL_KEY = "GLOBAL NUMBER IN MULTIPACK - ACTUAL"

# Fields the handbook marks "Sempre em branco" / "Sempre: X" - always the
# same fixed value regardless of description. These are auto-filled and
# locked; the hint popup explains why.
ALWAYS_BLANK_FIELDS = {
    "#BR LOC 100026 : ITEM EXCLUIDOS DOS FMT RETAIL/SCAN",
    "#BR LOC 102625 : ITEM DUMMY G2G",
    "#BR LOC 900901 : C-DAR CRITICOS",
    "#BR LOC 900900 : C-DAR DESCONTINUADOS",
    "#BR LOC 90020 : PROMOCAO",
    "#BR LOC 5000 : TOTAL MULTIPACK WITHOUT UNITS",
}
FIXED_VALUE_FIELDS = {
    ITM_CLASE_KEY: "CHINELOS",
    EMBALAGEM_KEY: "NAO DETERMINADO",
}

# Seed brand -> manufacturer knowledge (from the handbook + sample allocation).
# The tool ALSO learns new pairs automatically from any row the coder already
# completed in the same workbook (see App.learn_from_allocation()).
SEED_BRAND_MANUFACTURER = {
    "HAVAIANAS": "ALPARGATAS",
    "CARTAGO": "GRENDENE",
    "GRENDENE": "GRENDENE",
    "RIDER": "GRENDENE",
    "IPANEMA": "GRENDENE",
}

# Licensed-character / "Personagens" keywords -> VERSAO = PERSON (any brand).
PERSON_KEYWORDS = [
    "MICKEY","MINNIE","DISNEY","SNOOPY","PEANUTS","BELA E A FERA","BEAUTY AND THE BEAST",
    "PRINCESA","PRINCESS","TURMA DA M\u00d3NICA","TURMA DA MONICA","SONIC","HELLO KITTY",
    "PATRULHA CANINA","PAW PATROL","FROZEN","SPIDER","HOMEM ARANHA","BATMAN","SUPERMAN",
    "BARBIE","POKEMON","STAR WARS","MARVEL","AVENGERS","VINGADORES",
]
# HAVAIANAS-only version keywords (checked in this priority order).
HAVAIANAS_VERSION_KEYWORDS = [
    ("TOP", ["TOP"]),
    ("BRASIL", ["BRASIL", "BRAZIL"]),
    ("CASUAL", ["CASUAL"]),
    ("IPE", ["IPE", "IP\u00ca"]),
    ("TREND", ["TREND"]),
]
COUNTRY_FLAG_WORDS = ["ARGENTINA","ESTADOS UNIDOS","USA","FRANCA","FRANCE","ITALIA","ITALY",
                       "PORTUGAL","ESPANHA","SPAIN","ALEMANHA","GERMANY","JAPAO","JAPAN"]

# Strap format (FORMATO) - ESPECIAL keywords, else COMUM.
SPECIAL_STRAP_KEYWORDS = ["TRAN\u00c7AD", "TRANCAD", "CRUZAD", "TRENZAD", "BRAIDED", "CROSSED"]
# Strap length (COMPRIMENTO) - CURTA keywords, else LONGA (handbook default LONGA).
SHORT_STRAP_KEYWORDS = ["CURTA", "SHORT"]
# Strap thickness (ESPESSURA).
THIN_STRAP_KEYWORDS = ["FINA", "THIN"]
EXTRA_THICK_KEYWORDS = ["SUPER GROSSA", "EXTRA THICK", "EXTRA GROSSA"]
# Pattern / print (COLORA\u00c7\u00c3O) - ESTAMPADA keywords (includes degrad\u00ea/ombr\u00e9), else LISA.
PRINTED_KEYWORDS = ["ESTAMPAD", "PRINT", "DEGRAD", "OMBR\u00c9", "FLORAL", "LISTRAD", "STRIPED",
                     "ANIMAL PRINT", "ONCA", "ON\u00c7A", "TIE DYE"]

# Literal color lexicon (PT + EN) -> canonical uppercase value for the
# internal color field (#BR LOC 2025 : INTERNA CAMPO MT), matching the sample
# allocation data (AZUL, PRETO, MARROM, etc.).
COLOR_LEXICON = {
    "AZUL MARINHO": "AZUL", "NAVY": "AZUL", "AZUL": "AZUL", "BLUE": "AZUL",
    "PRETO": "PRETO", "BLACK": "PRETO",
    "BRANCO": "BRANCO", "WHITE": "BRANCO",
    "MARROM": "MARROM", "BROWN": "MARROM",
    "VERMELHO": "VERMELHO", "RED": "VERMELHO",
    "VERDE": "VERDE", "GREEN": "VERDE",
    "AMARELO": "AMARELO", "YELLOW": "AMARELO",
    "ROSA": "ROSA", "PINK": "ROSA",
    "CINZA": "CINZA", "GRAY": "CINZA", "GREY": "CINZA",
    "ROXO": "ROXO", "PURPLE": "ROXO", "LILAS": "ROXO", "LIL\u00c1S": "ROXO",
    "LARANJA": "LARANJA", "ORANGE": "LARANJA",
    "DOURADO": "DOURADO", "GOLD": "DOURADO",
    "PRATEADO": "PRATEADO", "SILVER": "PRATEADO",
    "BEGE": "BEGE", "BEIGE": "BEGE",
}

SIZE_PAIRS = ["17/18","19/20","21/22","23/24","25/26","27/28","29/30","31/32",
              "33/34","35/36","37/38","39/40","41/42","43/44"]

@dataclass
class Context:
    desc_pt: str            # original description, uppercase
    desc_en: str            # translated description, uppercase (falls back to desc_pt)
    tokens: set             # set of alnum tokens from desc_pt + desc_en
    proposed: dict          # field name -> current proposed value (may include values
                             # already picked earlier in this same auto-fill pass)
    known_brands: list      # brand names learned from the workbook, longest first
    brand_manufacturer: dict  # learned brand -> manufacturer map

    def has_any(self, keywords):
        blob = f" {self.desc_pt} {self.desc_en} "
        return any(k in blob for k in keywords)

def _blob(ctx: Context) -> str:
    return f" {ctx.desc_pt} {ctx.desc_en} "

def detect_brand(ctx: Context) -> str:
    blob = _blob(ctx)
    for brand in ctx.known_brands:
        if brand and f" {brand} " in blob:
            return brand
    return ""

def suggest_brand(ctx: Context):
    return detect_brand(ctx) or None

def suggest_manufacturer(ctx: Context):
    # Read the plain brand (MARCA), never BRAND 1 - BRAND 1 already contains
    # "BRAND (MANUFACTURER)" once populated, which would break the dict lookup.
    brand = ctx.proposed.get(BRAND_KEY) or detect_brand(ctx)
    if not brand:
        return None
    manufacturer = ctx.brand_manufacturer.get(brand) or SEED_BRAND_MANUFACTURER.get(brand)
    return manufacturer or None

def suggest_brand1(ctx: Context):
    """BOI/B1 parenthesis rule: BRAND (MANUFACTURER) unless they are identical."""
    brand = ctx.proposed.get(BRAND_KEY) or detect_brand(ctx)
    if not brand:
        return None
    manufacturer = ctx.proposed.get(MANUFACTURER_KEY) or suggest_manufacturer(ctx)
    if manufacturer and manufacturer != brand:
        return f"{brand} ({manufacturer})"
    return brand

STOP_WORDS = {"CHINELO","CHINELOS","SANDALIA","SANDALIAS","INFANTIL","FEMININO","MASCULINO",
              "ADULTO","UNISSEX","NA","DO","DA","DE","AO","E","PAR","TAMANHO","TAM","NUMERO",
              "N","DEDO","RASTEIRA"}

def suggest_subbrand(ctx: Context):
    """Model/line name left over after removing brand, manufacturer, category noise,
    size and detected color/pattern words. Falls back to the brand (handbook rule:
    'When it doesn't have any subbrand, populate with the brand value')."""
    brand = ctx.proposed.get(BRAND_KEY) or detect_brand(ctx)
    manufacturer = ctx.proposed.get(MANUFACTURER_KEY) or ""
    words = re.findall(r"[A-Z0-9\u00c0-\u00dc]+", ctx.desc_pt)
    leftover = []
    skip = STOP_WORDS | ({brand} if brand else set()) | ({manufacturer} if manufacturer else set())
    for w in words:
        if w in skip:
            continue
        if re.fullmatch(r"\d{1,3}", w):
            continue  # sizes handled separately
        if w in COLOR_LEXICON:
            continue
        leftover.append(w)
    if leftover:
        return " ".join(leftover[:3])
    return brand or None

def suggest_module(ctx: Context):
    if ctx.has_any(["INFANTIL", "BABY", "INFANT"]) and ctx.has_any(["ROUPA", "CLOTHING"]):
        return "CLOTHING BABY/INFANT"
    return "FOOTWEAR"

def suggest_formato(ctx: Context):
    return "ESPECIAL" if ctx.has_any(SPECIAL_STRAP_KEYWORDS) else "COMUM"

def suggest_comprimento(ctx: Context):
    return "CURTA" if ctx.has_any(SHORT_STRAP_KEYWORDS) else "LONGA"

def suggest_espessura(ctx: Context):
    if ctx.has_any(EXTRA_THICK_KEYWORDS):
        return "SUPER GROSSA"
    if ctx.has_any(THIN_STRAP_KEYWORDS):
        return "FINA"
    return "NORMAL"

def suggest_coloracao(ctx: Context):
    return "ESTAMPADA" if ctx.has_any(PRINTED_KEYWORDS) else "LISA"

def suggest_interna_mt(ctx: Context):
    blob = _blob(ctx)
    for phrase, canonical in COLOR_LEXICON.items():
        if f" {phrase} " in blob:
            return canonical
    return None

def suggest_versao(ctx: Context):
    if ctx.has_any(PERSON_KEYWORDS):
        return "PERSON"
    brand = ctx.proposed.get(BRAND_KEY) or detect_brand(ctx)
    if brand == "HAVAIANAS":
        for value, keywords in HAVAIANAS_VERSION_KEYWORDS:
            if ctx.has_any(keywords):
                return value
        if ctx.has_any(COUNTRY_FLAG_WORDS):
            return "PAISES"
        return "S/V"  # default - do not guess TRAD/TRADICIONAL without visual base-color check
    return "S/V"

def suggest_tamanho(ctx: Context):
    blob = _blob(ctx)
    pair = re.search(r"\b(\d{2})\s*[/\-]\s*(\d{2})\b", blob)
    if pair:
        return f"{pair.group(1)}/{pair.group(2)}"
    single = re.search(r"\b(?:TAM(?:ANHO)?\.?\s*)?(\d{2})\b", blob)
    if single:
        n = int(single.group(1))
        for p in SIZE_PAIRS:
            lo, hi = (int(x) for x in p.split("/"))
            if n in (lo, hi):
                return p
    return None

def suggest_itm_desc(ctx: Context):
    """Assistive concatenation per handbook: Brand, (Manufacturer via B1),
    Sub-brand (if different from brand), Version (if not S/V), Strap format
    (if ESPECIAL), Length (if CURTA), Thickness (if not NORMAL), Pattern
    (if ESTAMPADA), Size. Never repeats a value, keeps <=60 chars per the
    Learning Priority quality rules, and always needs a human review."""
    parts = []
    brand1 = ctx.proposed.get(BRAND1_KEY) or suggest_brand1(ctx)
    subbrand = ctx.proposed.get(SUBBRAND_KEY)
    versao = ctx.proposed.get(VERSAO_KEY)
    formato = ctx.proposed.get(FORMATO_KEY)
    comprimento = ctx.proposed.get(COMPRIMENTO_KEY)
    espessura = ctx.proposed.get(ESPESSURA_KEY)
    coloracao = ctx.proposed.get(COLORACAO_KEY)
    tamanho = ctx.proposed.get(TAMANHO_KEY)
    if brand1:
        parts.append(brand1)
    if subbrand and subbrand != (ctx.proposed.get(BRAND_KEY) or ""):
        parts.append(subbrand)
    if versao and versao != "S/V":
        parts.append(versao)
    if formato == "ESPECIAL":
        parts.append("ESPECIAL")
    if comprimento == "CURTA":
        parts.append("CURTA")
    if espessura and espessura != "NORMAL":
        parts.append(espessura)
    if coloracao == "ESTAMPADA":
        parts.append("ESTAMPADA")
    if tamanho:
        parts.append(tamanho)
    if not parts:
        return None
    result = " ".join(dict.fromkeys(parts))  # de-duplicate, keep order
    if len(result) > 60:
        result = result[:60].rstrip()
    return result

def suggest_ft_conv(ctx: Context):
    contenido = ctx.proposed.get(ITM_CONTENIDO_KEY)
    try:
        return str(int(float(contenido)) * 1000) if contenido else "1000"
    except Exception:
        return "1000"

PROMO_KEYWORDS = ["KIT", "LEVE", "PROMOCAO", "PROMOÇÃO", "GRATIS", "GR\u00c1TIS", "BRINDE", "COMBO"]

def suggest_promo_offer(ctx: Context):
    return "PRICE PROMOTION" if ctx.has_any(PROMO_KEYWORDS) else "NO PROMOTION"

def suggest_promo_type(ctx: Context):
    offer = ctx.proposed.get(PROMO_OFFER_KEY) or suggest_promo_offer(ctx)
    return "REGULAR" if offer == "NO PROMOTION" else None  # ambiguous -> needs manual review

# Master dispatch table: field name -> suggestion function.
SUGGEST_RULES = {
    MODULE_KEY: suggest_module,
    BRAND_KEY: suggest_brand,
    BRAND_G2G_KEY: suggest_brand,
    BRAND1_KEY: suggest_brand1,
    SUBBRAND_KEY: suggest_subbrand,
    MANUFACTURER_KEY: suggest_manufacturer,
    MANUFACTURER_G2G_KEY: suggest_manufacturer,
    BOI_KEY: suggest_manufacturer,
    FORMATO_KEY: suggest_formato,
    COMPRIMENTO_KEY: suggest_comprimento,
    ESPESSURA_KEY: suggest_espessura,
    COLORACAO_KEY: suggest_coloracao,
    INTERNA_MT_KEY: suggest_interna_mt,
    VERSAO_KEY: suggest_versao,
    TAMANHO_KEY: suggest_tamanho,
    ITEM_DESC_CHAR: suggest_itm_desc,
    FT_CONV_KEY: suggest_ft_conv,
    PROMO_OFFER_KEY: suggest_promo_offer,
    PROMO_TYPE_KEY: suggest_promo_type,
    ITM_CONTENIDO_KEY: lambda ctx: "1",
    FORMATO_TRANSF_KEY: lambda ctx: "VARIAVEL",
    GTI_PACK_KEY: lambda ctx: "NOT DETERMINED",
    GN_EACH_BASE_KEY: lambda ctx: "1'S",
    GN_EACH_ACTUAL_KEY: lambda ctx: "NOT STATED",
    GTP_MULTI_KEY: lambda ctx: "1 MU",
    GN_MULTI_BASE_KEY: lambda ctx: "X1",
    GN_MULTI_ACTUAL_KEY: lambda ctx: "X1",
}
for _name, _value in FIXED_VALUE_FIELDS.items():
    SUGGEST_RULES.setdefault(_name, (lambda v: (lambda ctx: v))(_value))
for _name in ALWAYS_BLANK_FIELDS:
    SUGGEST_RULES.setdefault(_name, lambda ctx: "")

def build_context(desc_pt, desc_en, proposed, known_brands, brand_manufacturer) -> Context:
    tokens = set(re.findall(r"[A-Z0-9\u00c0-\u00dc]+", f"{desc_pt} {desc_en}"))
    return Context(upper(desc_pt), upper(desc_en) or upper(desc_pt), tokens, dict(proposed),
                   known_brands, brand_manufacturer)

# ---------------------------------------------------------------------------
# Similarity detection - finds other items that are almost certainly the same
# product in a different size (or a very close color variant), so a value
# confirmed once can be copied across the whole size/color run instead of
# re-typing it item by item. This is the single biggest time-saver for a
# footwear catalog, where the same shoe is usually listed a dozen times.
# ---------------------------------------------------------------------------
_SIZE_TOKEN_RE = re.compile(r"^\d{2}(/\d{2})?$")

def _significant_tokens(desc:str)->set:
    """Tokens from a description with sizes and generic category/gender noise
    words stripped out, so similarity is judged on brand+model, not on the
    size or common words every item shares."""
    tokens = re.findall(r"[A-Z0-9\u00c0-\u00dc]+", upper(desc))
    return {t for t in tokens if t not in STOP_WORDS and not _SIZE_TOKEN_RE.match(t) and len(t) > 1}

def similarity_score(desc_a:str, desc_b:str)->float:
    """Jaccard similarity of the two descriptions' significant tokens (0-1)."""
    a, b = _significant_tokens(desc_a), _significant_tokens(desc_b)
    if not a or not b:
        return 0.0
    return len(a & b) / len(a | b)

# ---------------------------------------------------------------------------
# "Ask the Handbook" chatbot - two modes:
#
#  OFFLINE MODE (always available, zero setup): scores every guideline's
#  name/definitions/examples/extra notes/enumerated values, plus every
#  scope/worked-example note mined from the Definition and EXAMPLES sheets,
#  against the coder's question by keyword overlap (with PT<->EN synonym
#  expansion so a Portuguese question matches an English definition and vice
#  versa), and synthesizes a plain-language answer from the best matches.
#
#  AI MODE (optional, real-time): if the coder configures their own Anthropic
#  API key (Settings inside the chat window), questions are answered by a
#  real Claude model, given the entire handbook as context - genuinely
#  conversational, can handle follow-ups and phrasing offline search would
#  miss. Falls back to offline mode automatically if the call fails for any
#  reason (no key, bad key, no internet, bad model name), so the assistant
#  is never simply "down".
# ---------------------------------------------------------------------------
CHAT_STOPWORDS = {"O","A","OS","AS","DE","DA","DO","DAS","DOS","E","QUE","EM","UM","UMA","PARA","COM","THE",
                   "IS","OF","TO","AND","A","IN","WHAT","HOW","DOES","DO","I","WHEN","WHERE","SHOULD","VALUE",
                   "CHAR","CHARACTERISTIC","POPULATE","FILL","IT","THIS","THAT","ARE","CAN","WHICH"}

# PT <-> EN synonym pairs for the terms coders actually type, so a question in
# either language reliably matches guidelines written in the other.
CHAT_SYNONYMS = {
    "PREENCHER":{"FILL","POPULATE"},"COR":{"COLOR","COLOUR","COLORACAO"},"TAMANHO":{"SIZE"},
    "MARCA":{"BRAND"},"FABRICANTE":{"MANUFACTURER"},"TIRA":{"STRAP"},"ESPESSURA":{"THICKNESS"},
    "COMPRIMENTO":{"LENGTH"},"FORMATO":{"FORMAT","SHAPE"},"VERSAO":{"VERSION"},"ESTAMPADA":{"PRINTED"},
    "LISA":{"SOLID"},"PERSONAGEM":{"PERSON","CHARACTER","LICENSED"},"PERSONAGENS":{"PERSON","CHARACTER","LICENSED"},
    "PESO":{"WEIGHT"},"EMBALAGEM":{"PACKAGING","PACK"},"PROMOCAO":{"PROMOTION","PROMOTIONAL"},
    "SUBMARCA":{"SUBBRAND"},"CURTA":{"SHORT"},"LONGA":{"LONG"},"FINA":{"THIN"},"GROSSA":{"THICK"},
}

def _tokenize(value:str)->set:
    base={w for w in re.findall(r"[A-Z0-9\u00c0-\u00dc/]+",upper(value)) if w not in CHAT_STOPWORDS and len(w)>1}
    expanded=set(base)
    for tok in base:
        expanded |= CHAT_SYNONYMS.get(tok,set())
        for pt,en_set in CHAT_SYNONYMS.items():
            if tok in en_set:expanded.add(pt)
    return expanded

def _guideline_corpus(g:"Guideline")->str:
    parts=[g.name,g.pt_def,g.en_def,g.pt_examples,g.en_examples,g.extra_notes," ".join(g.enum_values)]
    if g.meta:parts.append(" ".join(g.meta.values()))
    return " ".join(p for p in parts if p)

def search_handbook(query:str, handbook:"HandbookData", top_n:int=3):
    """Return up to top_n (score, heading, body) tuples best matching the query,
    searched across the full enriched handbook corpus."""
    q_tokens=_tokenize(query)
    if not q_tokens:
        return []
    candidates=[]
    for g in handbook.guidelines.values():
        c_tokens=_tokenize(_guideline_corpus(g))
        score=len(q_tokens & c_tokens)
        if q_tokens & _tokenize(g.name):
            score+=4  # a hit on the characteristic's own name is a strong signal
        if score>0:
            body=f"Fixed/enumerated value? {g.fixed or 'N'}\n\nEN: {g.en_def or '-'}"
            if g.pt_def:body+=f"\n\nPT: {g.pt_def}"
            if g.enum_values:body+=f"\n\nValid values: {', '.join(g.enum_values)}"
            elif g.en_examples:body+=f"\n\nExample: {g.en_examples}"
            if g.extra_notes:body+=f"\n\nWorked examples from the handbook:\n{g.extra_notes}"
            if g.meta.get("country"):body+=f"\n\n(Scope: {g.meta.get('country')} / {g.meta.get('category','')})"
            candidates.append((score,g.name,body))
    for note in handbook.scope_notes:
        score=len(q_tokens & _tokenize(note))
        if score>0:
            candidates.append((score,"Handbook note",note))
    candidates.sort(key=lambda c:-c[0])
    # de-duplicate near-identical bodies (the same note can be mined twice)
    seen=set(); unique=[]
    for c in candidates:
        sig=c[2][:80]
        if sig in seen:continue
        seen.add(sig); unique.append(c)
    return unique[:top_n]

def synthesize_offline_answer(query:str, handbook:"HandbookData")->str:
    hits=search_handbook(query,handbook,top_n=3)
    if not hits:
        return ("I couldn't find anything in the handbook matching that question. Try naming the "
                "characteristic directly (e.g. 'FORMATO', 'TAMANHO', 'VERSAO') or a keyword from its rule.")
    return "\n\n".join(f"\U0001f4cc {h[1]}\n{h[2]}" for h in hits)

# -- Optional real-time AI mode (Anthropic API, using the coder's own key) --
ANTHROPIC_API_URL = "https://api.anthropic.com/v1/messages"
DEFAULT_AI_MODEL = "claude-sonnet-5"

def build_ai_system_prompt(handbook:"HandbookData", item_context:str="")->str:
    lines=["You are a coding-guideline assistant helping a Brazil OGRDS Item Coding team member "
           "for LPC 23700 (Chinelos / flip-flops). Answer ONLY using the guideline knowledge below. "
           "Be concise, practical, and point to the exact characteristic name and valid values when relevant. "
           "If something isn't covered by the guidelines below, say so plainly instead of guessing."]
    if item_context:
        lines.append(f"\nThe coder is currently working on this item: {item_context}")
    lines.append("\n=== CHARACTERISTIC GUIDELINES ===")
    for g in handbook.guidelines.values():
        entry=f"\n- {g.name} (fixed/enumerated: {g.fixed or 'N'})\n  EN: {g.en_def[:400]}"
        if g.enum_values:entry+=f"\n  Valid values: {', '.join(g.enum_values)}"
        if g.extra_notes:entry+=f"\n  Worked examples: {g.extra_notes[:400]}"
        lines.append(entry)
    if handbook.scope_notes:
        lines.append("\n=== SCOPE / ADDITIONAL NOTES ===")
        for note in handbook.scope_notes[:25]:
            lines.append(f"- {note[:300]}")
    return "\n".join(lines)

def call_claude_chatbot(question:str, handbook:"HandbookData", api_key:str, model:str,
                         history:list, item_context:str="")->str:
    """Calls the real Anthropic Messages API using the coder's own key. Raises
    on any failure - the caller is expected to catch and fall back to the
    offline engine, so a bad key/model/network never looks like a crash."""
    import urllib.request, urllib.error
    system_prompt=build_ai_system_prompt(handbook,item_context)
    messages=history+[{"role":"user","content":question}]
    payload=json.dumps({"model":model or DEFAULT_AI_MODEL,"max_tokens":700,
                         "system":system_prompt,"messages":messages}).encode("utf-8")
    req=urllib.request.Request(ANTHROPIC_API_URL,data=payload,method="POST",headers={
        "content-type":"application/json","x-api-key":api_key,"anthropic-version":"2023-06-01"})
    with urllib.request.urlopen(req,timeout=25) as resp:
        data=json.loads(resp.read().decode("utf-8"))
    parts=[b.get("text","") for b in data.get("content",[]) if b.get("type")=="text"]
    reply="".join(parts).strip()
    if not reply:
        raise RuntimeError("Empty response from the API")
    return reply

# ---------------------------------------------------------------------------
# Main application
# ---------------------------------------------------------------------------
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("light"); ctk.set_default_color_theme("blue")
        self.title(APP); self.configure(fg_color="#F5F8FC"); self.minsize(1150,700)
        self.geometry(f"{min(1750,self.winfo_screenwidth()-25)}x{min(980,self.winfo_screenheight()-60)}+6+6")
        self.wb=self.ws=None; self.path=None; self.fields: list[Field]=[]
        self.handbook = HandbookData()
        self.groups: dict[str,list[int]] = {}
        self.filtered: list[str] = []; self.position = 0
        self.vars={}; self.boxes={}; self.hint_buttons={}; self.suggest_buttons={}
        self.lock_buttons={}; self.locked_fields=set(); self.locks={}
        self.session_start=datetime.now(); self.session_done_count=0
        self.known_brands: list[str] = []
        self.brand_manufacturer: dict[str,str] = {}
        self.known_values: dict[str,set] = defaultdict(set)
        self.loading = False; self.dirty = False
        self.settings_file = Path.home()/".brazil_item_coding_tool.json"
        self.build()

    # -- layout ------------------------------------------------------------
    def build(self):
        header=ctk.CTkFrame(self,fg_color="#0B7A45",corner_radius=0,height=62); header.pack(fill="x"); header.pack_propagate(False)
        branding=ctk.CTkFrame(header,fg_color="transparent"); branding.pack(side="left",padx=18,pady=5)
        ctk.CTkLabel(branding,text="Brazil Item Coding Studio",font=ctk.CTkFont(size=21,weight="bold"),text_color="white",anchor="w").pack(anchor="w")
        ctk.CTkLabel(branding,text="All-in-One Version 2.0  |  LPC 23700 - Chinelos  |  Created by Sumit Mondal",
                    font=ctk.CTkFont(size=10),text_color="#D8F6E4",anchor="w").pack(anchor="w")
        self.file_buttons={}
        header_buttons=ctk.CTkFrame(header,fg_color="transparent"); header_buttons.pack(side="right",padx=6,pady=9)
        ctk.CTkButton(header_buttons,text="\U0001f50d Quality Audit",command=self.run_quality_audit,width=150,height=30,
                     fg_color="#B0791F",hover_color="#93630F").pack(side="right",padx=4)
        ctk.CTkButton(header_buttons,text="\U0001f4ac Ask the Handbook",command=self.open_chatbot,width=170,height=30,
                     fg_color="#7A3FD8",hover_color="#622FB0").pack(side="right",padx=4)
        for label,command,key,color in [("Allocation",self.browse_allocation,"allocation","white"),
                                         ("Handbook",self.browse_handbook,"handbook","#2FA36B"),
                                         ("IDE / Reference",self.browse_ide,"ide","#1F7A9D")]:
            button=ctk.CTkButton(header_buttons,text="Load "+label,command=command,width=150,height=30,
                                 fg_color=color,text_color="#0B7A45" if color=="white" else "white")
            button.pack(side="right",padx=4); self.file_buttons[key]=button

        search=ctk.CTkFrame(self,fg_color="white",border_width=1,border_color="#D3E0EA"); search.pack(fill="x",padx=12,pady=6)
        self.search_var=StringVar()
        search_entry=ctk.CTkEntry(search,textvariable=self.search_var,placeholder_text="Search description, barcode or lpc id",height=32)
        search_entry.pack(side="left",fill="x",expand=True,padx=8,pady=5); search_entry.bind("<KeyRelease>",lambda e:self.apply_filter())
        self.filter_menu=ctk.CTkOptionMenu(search,values=["ALL","DONE","PENDING","QUERY","OUT OF SCOPE"],width=140,height=32,command=lambda x:self.apply_filter())
        self.filter_menu.set("ALL"); self.filter_menu.pack(side="left",padx=4)
        self.status=ctk.CTkLabel(search,text="No workbook loaded",font=ctk.CTkFont(size=12,weight="bold"),text_color="#0B7A45"); self.status.pack(side="right",padx=10)

        # -- progress dashboard --------------------------------------------
        dash=ctk.CTkFrame(self,fg_color="white",border_width=1,border_color="#D3E0EA"); dash.pack(fill="x",padx=12,pady=(0,6))
        self.progress_bar=ctk.CTkProgressBar(dash,height=14,progress_color="#188A49"); self.progress_bar.set(0)
        self.progress_bar.pack(side="left",fill="x",expand=True,padx=(10,10),pady=8)
        self.progress_label=ctk.CTkLabel(dash,text="No items loaded yet",font=ctk.CTkFont(size=10,weight="bold"),text_color="#333")
        self.progress_label.pack(side="left",padx=(0,10))
        ctk.CTkButton(dash,text="\U0001f4c4 Export Summary Report",command=self.export_summary_report,width=190,height=26,
                     fg_color="#5A6472",hover_color="#454E59").pack(side="right",padx=8,pady=6)

        # -- item info + translation panel --
        self.info=ctk.CTkFrame(self,fg_color="#EAF7EF",border_width=1,border_color="#BFE3CC",corner_radius=6); self.info.pack(fill="x",padx=12,pady=2)
        row1=ctk.CTkFrame(self.info,fg_color="transparent"); row1.pack(fill="x",padx=8,pady=(6,2))
        ctk.CTkLabel(row1,text="EXTERNAL CODE:",font=ctk.CTkFont(size=9,weight="bold"),text_color="#0B7A45").pack(side="left")
        self.external_code_lbl=ctk.CTkLabel(row1,text="",font=ctk.CTkFont(size=10)); self.external_code_lbl.pack(side="left",padx=(4,20))
        ctk.CTkLabel(row1,text="ITEM STATUS:",font=ctk.CTkFont(size=9,weight="bold"),text_color="#0B7A45").pack(side="left")
        self.item_status_var=StringVar(value="PENDING")
        ctk.CTkOptionMenu(row1,variable=self.item_status_var,values=["PENDING","DONE","QUERY","OUT OF SCOPE"],width=140,height=26,
                           command=self.item_status_changed).pack(side="left",padx=(4,0))

        row2=ctk.CTkFrame(self.info,fg_color="transparent"); row2.pack(fill="x",padx=8,pady=2)
        ctk.CTkLabel(row2,text="ORIGINAL DESCRIPTION:",font=ctk.CTkFont(size=9,weight="bold"),text_color="#0B7A45",width=170,anchor="w").pack(side="left")
        self.original_desc_lbl=ctk.CTkLabel(row2,text="",font=ctk.CTkFont(size=11),anchor="w",justify="left",wraplength=1100,cursor="hand2",text_color="#0B4D2C")
        self.original_desc_lbl.pack(side="left",fill="x",expand=True); self.original_desc_lbl.bind("<Button-1>",lambda e:self.search_description_online())

        row3=ctk.CTkFrame(self.info,fg_color="transparent"); row3.pack(fill="x",padx=8,pady=(2,6))
        ctk.CTkLabel(row3,text="ENGLISH TRANSLATION:",font=ctk.CTkFont(size=9,weight="bold"),text_color="#7A5A0B",width=170,anchor="w").pack(side="left")
        self.translated_desc_lbl=ctk.CTkLabel(row3,text="",font=ctk.CTkFont(size=11,slant="italic"),anchor="w",justify="left",wraplength=980,text_color="#7A5A0B")
        self.translated_desc_lbl.pack(side="left",fill="x",expand=True)
        self.translate_note_lbl=ctk.CTkLabel(row3,text="",font=ctk.CTkFont(size=8),text_color="#B08900"); self.translate_note_lbl.pack(side="right",padx=6)
        ctk.CTkButton(row3,text="Re-translate",command=self.refresh_translation,width=110,height=24,
                      fg_color="#E6B800",hover_color="#C79D00",text_color="#3A2E00").pack(side="right",padx=4)

        comment=ctk.CTkFrame(self.info,fg_color="transparent"); comment.pack(fill="x",padx=9,pady=(0,6))
        ctk.CTkLabel(comment,text="COMMENT:",width=95,anchor="w",font=ctk.CTkFont(size=9,weight="bold"),text_color="#0B7A45").pack(side="left")
        self.comment=ctk.CTkTextbox(comment,height=27,font=ctk.CTkFont(size=10),fg_color="#FFF7D6",border_width=1,border_color="#D7BD60")
        self.comment.pack(side="left",fill="x",expand=True); self.comment.bind("<KeyRelease>",lambda e:self.mark_dirty())

        titles=ctk.CTkFrame(self,fg_color="#DCF5E8",corner_radius=0); titles.pack(fill="x",padx=12,pady=(3,0))
        ctk.CTkLabel(titles,text="CHARACTERISTIC",width=330,anchor="w",font=ctk.CTkFont(size=12,weight="bold"),text_color="#0B4D2C").pack(side="left",padx=9,pady=6)
        ctk.CTkLabel(titles,text="HINT",width=220,anchor="w",font=ctk.CTkFont(size=11,weight="bold"),text_color="#176A65").pack(side="left",padx=5)
        ctk.CTkLabel(titles,text="PROPOSED VALUE",anchor="w",font=ctk.CTkFont(size=12,weight="bold"),text_color="#7B6410").pack(side="left",padx=5,fill="x",expand=True)

        # Custom ScrollableFrame (see class above) - fixes trackpad two-finger
        # scroll reliability that CTkScrollableFrame's own binding struggled with.
        self.scroll=ScrollableFrame(self,fg_color="#F5F8FC",corner_radius=0); self.scroll.pack(fill="both",expand=True,padx=12)

        footer=ctk.CTkFrame(self,fg_color="white",border_width=1,border_color="#D3E0EA"); footer.pack(fill="x",padx=12,pady=(3,8))
        nav=ctk.CTkFrame(footer,fg_color="transparent"); nav.pack(side="left",padx=7,pady=7)
        ctk.CTkButton(nav,text="\u25c0 Previous",command=self.previous,width=95,height=30).pack(side="left",padx=2)
        ctk.CTkButton(nav,text="Next \u25b6",command=self.next,width=95,height=30).pack(side="left",padx=2)
        ctk.CTkButton(nav,text="\U0001f4a1 Auto-fill",command=self.auto_fill_all,width=110,height=30,
                      fg_color="#2FA36B",hover_color="#238754").pack(side="left",padx=(12,2))
        ctk.CTkButton(nav,text="\U0001f517 Similar Items",command=self.open_similar_items_dialog,width=140,height=30,
                     fg_color="#2F80C1",hover_color="#26669C").pack(side="left",padx=2)
        ctk.CTkButton(nav,text="\U0001f512 Lock All Filled",command=self.lock_all_filled,width=140,height=30,
                     fg_color="#188A49",hover_color="#136B38").pack(side="left",padx=2)
        ctk.CTkButton(nav,text="\U0001f513 Unlock All",command=self.unlock_all,width=120,height=30,
                     fg_color="#8A8F98",hover_color="#6E727A").pack(side="left",padx=2)
        ctk.CTkButton(nav,text="\u267b Recover",command=self.recover_local_work,width=100,height=30,
                      fg_color="#B0791F",hover_color="#93630F").pack(side="left",padx=2)

        save_group=ctk.CTkFrame(footer,fg_color="transparent"); save_group.pack(side="right",padx=7,pady=7)
        ctk.CTkButton(save_group,text="Mark Done + Save",command=self.mark_done,width=150,height=30,fg_color="#188A49").pack(side="right",padx=2)
        ctk.CTkButton(save_group,text="\u2601 Save Changes Online",command=self.save_changes_online,width=175,height=30,
                     fg_color="#1F6FB2",hover_color="#175A90").pack(side="right",padx=2)
        ctk.CTkButton(save_group,text="Save Draft",command=lambda:self.save(),width=100,height=30).pack(side="right",padx=2)
        self.theme_switch=ctk.CTkSegmentedButton(footer,values=["Light","Dark"],width=110,height=26,command=self.toggle_theme)
        self.theme_switch.set("Light"); self.theme_switch.pack(side="right",padx=(0,14))
        ctk.CTkLabel(footer,text=f"{APP}  \u2022  Created by Sumit Mondal",font=ctk.CTkFont(size=9),text_color="#56778C").pack(side="left",padx=16)
        self.bind("<Control-s>",lambda e:self.save())
        self.bind("<Control-Right>",lambda e:self.next())
        self.bind("<Control-Left>",lambda e:self.previous())
        self.bind("<Control-Return>",lambda e:self.mark_done())
        self.after(300,self.auto_load_files)

    # -- settings / auto-load ----------------------------------------------
    def read_settings(self):
        try:return json.loads(self.settings_file.read_text(encoding="utf-8"))
        except Exception:return {}
    def remember_path(self,key,path):
        data=self.read_settings(); data[key]=str(Path(path).resolve())
        self.settings_file.write_text(json.dumps(data,indent=2),encoding="utf-8")
    def mark_file_linked(self,key,path):
        button=self.file_buttons.get(key)
        if button:
            short=Path(path).stem
            if len(short)>16:short=short[:13]+"..."
            button.configure(text="Linked \u2713 "+short,fg_color="#1F9D55",hover_color="#188247",text_color="white")
    def auto_load_files(self):
        settings=self.read_settings()
        for key,loader in [("handbook",self.load_handbook_file),("ide",self.load_ide_file),("allocation",self.load_allocation)]:
            saved=settings.get(key,"")
            if saved and Path(saved).exists():
                try:loader(Path(saved),quiet=True)
                except Exception as exc:
                    messagebox.showwarning(APP,f"Could not reload {key}: {exc}")

    def browse_allocation(self):
        p=filedialog.askopenfilename(filetypes=[("Excel","*.xlsx")])
        if p:
            try:self.load_allocation(Path(p))
            except Exception as exc:messagebox.showerror(APP,f"Allocation could not be loaded.\n\n{exc}")
    def browse_handbook(self):
        p=filedialog.askopenfilename(filetypes=[("Excel","*.xlsx")])
        if p:
            try:self.load_handbook_file(Path(p))
            except Exception as exc:messagebox.showerror(APP,f"Handbook could not be loaded.\n\n{exc}")
    def browse_ide(self):
        p=filedialog.askopenfilename(filetypes=[("Excel","*.xlsx")])
        if p:
            try:self.load_ide_file(Path(p))
            except Exception as exc:messagebox.showerror(APP,f"IDE / reference file could not be loaded.\n\n{exc}")

    def load_ide_file(self,path,quiet=False):
        """Load a reference / IDE workbook of already-coded items (any layout).
        Any column whose header matches one of our characteristic names feeds
        extra values into that field's dropdown; if both a brand-like and a
        manufacturer-like column are present, brand->manufacturer pairs are
        learned from it too - exactly like the India tool's IDE linking."""
        self.status.configure(text="Loading IDE / reference file...");self.update_idletasks()
        wb=load_workbook(path,read_only=True,data_only=True)
        ws=wb[wb.sheetnames[0]]
        rows=list(ws.iter_rows(values_only=True))
        wb.close()
        if not rows:
            raise ValueError("The reference file is empty.")
        header=[text(h) for h in rows[0]]
        field_names={f.name for f in self.fields} if self.fields else set()
        col_index={h:i for i,h in enumerate(header)}
        added=0
        for h,i in col_index.items():
            if field_names and h not in field_names:
                # also allow a plain "Proposed Value - X" header, matching the allocation convention
                if h.startswith(PROPOSED_PREFIX) and h[len(PROPOSED_PREFIX):] in field_names:
                    target=h[len(PROPOSED_PREFIX):]
                else:
                    continue
            else:
                target=h
            for r in rows[1:]:
                if i<len(r):
                    v=upper(r[i])
                    if v:
                        self.known_values[target].add(v); added+=1
        def find_col(*keys):
            for k in keys:
                if k in col_index:
                    return col_index[k]
            return None
        brand_col=find_col(BRAND_KEY,PROPOSED_PREFIX+BRAND_KEY)
        man_col=find_col(MANUFACTURER_KEY,PROPOSED_PREFIX+MANUFACTURER_KEY)
        learned_pairs=0
        if brand_col is not None and man_col is not None:
            for r in rows[1:]:
                b=upper(r[brand_col]) if brand_col<len(r) else ""
                m=upper(r[man_col]) if man_col<len(r) else ""
                if b and m and m not in {"NAO DETERMINADO","NOT DETERMINED","NO IDENTIFICADO","NAO IDENTIFICADO"}:
                    self.brand_manufacturer[b]=m; learned_pairs+=1
            self.known_brands=sorted(set(self.known_brands)|set(self.brand_manufacturer),key=lambda b:(-len(b),b))
        if self.fields:self.build_rows()
        self.remember_path("ide",path); self.mark_file_linked("ide",path)
        if not quiet:
            messagebox.showinfo(APP,f"IDE / reference file loaded: {added} value(s) learned, {learned_pairs} brand\u2192manufacturer pair(s) learned.")

    def load_handbook_file(self,path,quiet=False):
        self.status.configure(text="Loading handbook (guidelines + reference photos)...");self.update_idletasks()
        self.handbook=load_handbook(path)
        self.enrich_known_values_from_guidelines()
        self.remember_path("handbook",path); self.mark_file_linked("handbook",path)
        if self.fields:self.build_rows()
        if not quiet:
            messagebox.showinfo(APP,f"Handbook loaded: {len(self.handbook.guidelines)} characteristic guideline(s), "
                                     f"{sum(len(v) for v in self.handbook.images.values())} reference photo(s), "
                                     f"{len(self.handbook.scope_notes)} worked-example note(s)")

    def enrich_known_values_from_guidelines(self):
        """Seed each field's dropdown with the handbook's own enumerated valid
        values (e.g. VERSAO's S/V, PERSON, TOP, BRASIL...), so the full set of
        allowed values is available immediately - even before any item has
        ever had that field filled in."""
        if not hasattr(self,"known_values") or self.known_values is None:
            self.known_values=defaultdict(set)
        for name,g in self.handbook.guidelines.items():
            for v in g.enum_values:
                clean=re.sub(r"\s*\(.*?\)\s*","",v).strip()  # drop parenthetical notes like "(used for version Personagens)"
                if clean:self.known_values[name].add(clean)

    @staticmethod
    def parse_sheet(ws):
        """Parse an allocation worksheet into (fields, hmap, groups). Shared by
        load_allocation() and save_changes_online() so both always agree on
        how columns and item-groups are identified."""
        headers=[text(c.value) for c in ws[1]]
        fields=[]
        for c,h in enumerate(headers,1):
            if h.startswith(PROPOSED_PREFIX):
                base=h[len(PROPOSED_PREFIX):]
                source_col=c-1 if c-1>=1 and headers[c-2]==base else None
                fields.append(Field(base,source_col,c))
        if not fields:
            raise ValueError("No 'Proposed Value - <Characteristic>' columns were found in this allocation file.")
        hmap={h:i for i,h in enumerate(headers,1)}
        for wanted in WORKFLOW_COLUMNS:
            if wanted not in hmap:
                col=ws.max_column+1; ws.cell(1,col).value=wanted; hmap[wanted]=col
        desc_col=hmap.get(DESCRIPTION_HEADER)
        if not desc_col:
            raise ValueError(f"'{DESCRIPTION_HEADER}' column was not found.")
        groups=defaultdict(list)
        for r in range(2,ws.max_row+1):
            desc=text(ws.cell(r,desc_col).value)
            code=text(ws.cell(r,hmap.get("External Code",desc_col)).value)
            if not desc and not code:continue
            key=desc or f"CODE:{code}"
            groups[key].append(r)
        return fields,hmap,dict(groups)

    # -- allocation loading --------------------------------------------------
    def load_allocation(self,path,quiet=False):
        if self.wb:
            self.save(silent=True); self.wb.close()
        self.status.configure(text="Loading allocation workbook..."); self.update_idletasks()
        self.path=Path(path); self.wb=load_workbook(path); self.ws=self.wb[self.wb.sheetnames[0]]
        self.fields,self.hmap,self.groups=self.parse_sheet(self.ws)
        self.learn_from_allocation()
        self.enrich_known_values_from_guidelines()
        self.load_locks()
        self.build_rows()
        self.remember_path("allocation",path); self.mark_file_linked("allocation",path)
        self.filtered=list(self.groups.keys()); self.position=0
        self.apply_filter(reset=True)
        self.update_progress_dashboard()
        self.status.configure(text=f"Loaded {len(self.groups):,} item(s) | {len(self.fields)} characteristics")

    def learn_from_allocation(self):
        """Self-learning brand/manufacturer knowledge from whatever this workbook
        already has proposed (mirrors the India tool's IDE-based learning).
        This MERGES into (never replaces) whatever the IDE/reference file
        already taught the tool, regardless of load order."""
        if not hasattr(self,"known_values") or self.known_values is None:
            self.known_values=defaultdict(set)
        if not hasattr(self,"brand_manufacturer") or self.brand_manufacturer is None:
            self.brand_manufacturer={}
        field_by_name={f.name:f for f in self.fields}
        brand_field=field_by_name.get(BRAND_KEY); man_field=field_by_name.get(MANUFACTURER_KEY)
        for f in self.fields:
            col=f.proposed_col
            for rows in self.groups.values():
                for r in rows:
                    v=upper(self.ws.cell(r,col).value)
                    if v:self.known_values[f.name].add(v)
        if brand_field and man_field:
            for rows in self.groups.values():
                r=rows[0]
                b=upper(self.ws.cell(r,brand_field.proposed_col).value)
                m=upper(self.ws.cell(r,man_field.proposed_col).value)
                if b and m and m not in {"NAO DETERMINADO","NOT DETERMINED","NO IDENTIFICADO","NAO IDENTIFICADO"}:
                    self.brand_manufacturer[b]=m
        seed_brands=set(SEED_BRAND_MANUFACTURER)|set(self.brand_manufacturer)|self.known_values.get(BRAND_KEY,set())
        self.known_brands=sorted((b for b in seed_brands if b and b not in {"NAO DETERMINADO","NOT DETERMINED","NO IDENTIFICADO","NAO IDENTIFICADO"}),
                                  key=lambda b:(-len(b),b))
        for k,v in SEED_BRAND_MANUFACTURER.items():
            self.brand_manufacturer.setdefault(k,v)

    # -- row building --------------------------------------------------------
    def build_rows(self):
        for w in self.scroll.content.winfo_children():w.destroy()
        self.vars={};self.boxes={};self.hint_buttons={};self.suggest_buttons={};self.source_labels={};self.lock_buttons={};self.locked_fields=getattr(self,"locked_fields",set())
        for f in self.fields:
            row=ctk.CTkFrame(self.scroll.content,fg_color="#FDFEFF",border_width=1,border_color="#D9E6EF",height=46)
            row.pack(fill="x",pady=1); row.pack_propagate(False)  # fixed row height - stops big gaps between rows
            label_box=ctk.CTkFrame(row,fg_color="transparent",width=330); label_box.pack(side="left",padx=(9,0),pady=5); label_box.pack_propagate(False)
            ctk.CTkLabel(label_box,text=f.name,anchor="w",wraplength=290,justify="left",font=ctk.CTkFont(size=11,weight="bold")).pack(side="left",fill="x",expand=True)
            hint=ctk.CTkButton(label_box,text="?",width=24,height=24,corner_radius=12,fg_color="#3188C6",hover_color="#276A9C",
                               command=lambda n=f.name:self.open_hint(n))
            hint.pack(side="right",padx=(2,0)); self.hint_buttons[f.name]=hint

            source_lbl=ctk.CTkLabel(row,text="",width=220,anchor="w",wraplength=210,justify="left",font=ctk.CTkFont(size=10),text_color="#176A65")
            source_lbl.pack(side="left",padx=5,pady=5); self.source_labels[f.name]=source_lbl

            proposed_panel=ctk.CTkFrame(row,fg_color="transparent"); proposed_panel.pack(side="left",fill="both",expand=True,padx=(4,7),pady=2)
            var=StringVar()
            box=SearchEntry(proposed_panel,textvariable=var,height=32,values=self.known_values.get(f.name,set()),
                            fg_color="#FFF9DD",border_color="#D7BD60",on_pick=lambda v,n=f.name:self.on_value_picked(n))
            box.pack(side="left",fill="x",expand=True,padx=(0,5),pady=5)
            var.trace_add("write",lambda *a,n=f.name:self.mark_dirty())
            suggest=ctk.CTkButton(proposed_panel,text="\U0001f4a1",width=34,height=30,corner_radius=6,fg_color="#2FA36B",hover_color="#238754",
                                  command=lambda n=f.name:self.suggest_field(n,overwrite=True))
            suggest.pack(side="right",padx=(0,1),pady=5); self.suggest_buttons[f.name]=suggest
            lock_btn=ctk.CTkButton(proposed_panel,text="\U0001f513",width=34,height=30,corner_radius=6,fg_color="#8A8F98",hover_color="#6E727A",
                                   command=lambda n=f.name:self.toggle_lock(n))
            lock_btn.pack(side="right",padx=(0,4),pady=5); self.lock_buttons[f.name]=lock_btn
            self.vars[f.name]=var; self.boxes[f.name]=box

    def images_for_field(self,name):
        """Best-effort lookup of reference photos for a characteristic, matching
        on the handbook's '#BR LOC ... : CHAR' label prefix."""
        found=[]
        base=upper(name)
        for label,blobs in self.handbook.images.items():
            if base in label or label in base:
                found.extend(blobs)
        return found[:4]

    def open_hint(self,name):
        g=self.handbook.guidelines.get(name)
        win=Toplevel(self); win.title(f"Guideline - {name}"); win.geometry("700x680"); win.attributes("-topmost",True)
        outer=ctk.CTkScrollableFrame(win,fg_color="white"); outer.pack(fill="both",expand=True,padx=6,pady=6)
        text_box=Text(outer,wrap="word",font=("Segoe UI",10),padx=10,pady=10,height=18,relief="flat")
        text_box.pack(fill="x")
        if not g:
            text_box.insert("end",f"No guideline found for '{name}'.\n\nLoad the handbook workbook (Load Handbook) to see the "
                                   "PT/EN definition, whether the value is fixed, and worked examples for this characteristic.")
        else:
            text_box.insert("end", f"CHARACTERISTIC\n{g.name}\n\n")
            text_box.insert("end", f"VALUE FIXED?\n{g.fixed or 'N'}\n\n")
            text_box.insert("end", f"DEFINITION (EN)\n{g.en_def or '-'}\n\n")
            text_box.insert("end", f"DEFINIC\u00c3O (PT)\n{g.pt_def or '-'}\n\n")
            if g.en_examples:
                text_box.insert("end", f"EXAMPLES (EN)\n{g.en_examples}\n\n")
            if g.pt_examples:
                text_box.insert("end", f"EXEMPLOS (PT)\n{g.pt_examples}\n\n")
        suggestion=self.compute_suggestion(name)
        if suggestion is not None:
            text_box.insert("end", f"\nCURRENT AUTO-SUGGESTION FOR THIS ITEM\n{suggestion or '(blank - handbook says leave blank)'}\n")
        text_box.configure(state="disabled")

        photos=self.images_for_field(name)
        if photos:
            ctk.CTkLabel(outer,text="REFERENCE PHOTOS FROM THE HANDBOOK",font=ctk.CTkFont(size=11,weight="bold"),
                        text_color="#0B7A45").pack(anchor="w",padx=6,pady=(10,4))
            gallery=ctk.CTkFrame(outer,fg_color="transparent"); gallery.pack(fill="x",padx=6)
            win._photo_refs=[]  # keep references alive for the life of the popup
            for blob in photos:
                try:
                    img=PhotoImage(data=blob)
                    while img.width()>220 or img.height()>220:
                        img=img.subsample(2,2)
                    win._photo_refs.append(img)
                    tk.Label(gallery,image=img,bd=1,relief="solid").pack(side="left",padx=4,pady=4)
                except Exception:
                    continue
        else:
            ctk.CTkLabel(outer,text="(No matching reference photo found in the handbook for this characteristic.)",
                        font=ctk.CTkFont(size=9,slant="italic"),text_color="#888").pack(anchor="w",padx=6,pady=(10,4))
        ctk.CTkButton(win,text="Close",command=win.destroy).pack(pady=6)

    # -- translation ---------------------------------------------------------
    def refresh_translation(self):
        desc=self.original_desc_lbl.cget("text")
        if desc in _TRANSLATE_CACHE:
            del _TRANSLATE_CACHE[desc]  # force a fresh attempt on manual re-translate
        translated,mode=translate_to_english(desc)
        self.translated_desc_lbl.configure(text=translated)
        self.translate_note_lbl.configure(text=mode)

    def search_description_online(self):
        desc=self.original_desc_lbl.cget("text")
        if desc:webbrowser.open("https://www.google.com/search?q="+quote_plus(desc))

    # -- context / suggestions ------------------------------------------------
    def build_context(self):
        desc_pt=self.original_desc_lbl.cget("text")
        desc_en=self.translated_desc_lbl.cget("text")
        proposed={n:upper(v.get()) for n,v in self.vars.items() if upper(v.get())}
        return build_context(desc_pt,desc_en,proposed,self.known_brands,self.brand_manufacturer)

    def compute_suggestion(self,name):
        rule=SUGGEST_RULES.get(name)
        if not rule:return None
        ctx=self.build_context()
        try:return rule(ctx)
        except Exception:return None

    def suggest_field(self,name,overwrite=False):
        if name not in self.vars:return
        if name in self.locked_fields:return  # locked fields are never touched by suggestions
        if not overwrite and upper(self.vars[name].get()):return  # never clobber a filled value during bulk auto-fill
        value=self.compute_suggestion(name)
        if value is None:
            if overwrite:self.toast(f"No confident suggestion for {name} - please fill manually")
            return
        self.vars[name].set(value); self.mark_dirty()
        if overwrite:self.lock_field(name)  # accepting a single-field suggestion confirms and locks it

    def auto_fill_all(self):
        if not self.current_key():return
        # Two passes: fields with no dependency first (brand/manufacturer/etc.),
        # then ITM_DESC which is built from whatever the first pass produced.
        order=[f.name for f in self.fields if f.name != ITEM_DESC_CHAR] + [ITEM_DESC_CHAR]
        filled=0
        for name in order:
            if name not in self.vars:continue
            before=upper(self.vars[name].get())
            self.suggest_field(name,overwrite=False)
            if not before and upper(self.vars[name].get()):filled+=1
        self.toast(f"Auto-filled {filled} characteristic(s) - please review before saving")

    # -- lock / unlock --------------------------------------------------------
    def on_value_picked(self,name):
        """Called whenever the coder confirms a value - by picking from the
        dropdown, pressing Enter/Tab, or clicking away from the field. This is
        the 'I'm done with this field' signal, so the value is locked right
        away to protect it from being accidentally overwritten; Unlock is
        always one click away if a correction is needed."""
        self.mark_dirty()
        self.lock_field(name)

    def lock_field(self,name):
        if not upper(self.vars.get(name,StringVar()).get()):return
        self.locked_fields.add(name); self.apply_lock_visual(name)
    def unlock_field(self,name):
        self.locked_fields.discard(name); self.apply_lock_visual(name)
    def toggle_lock(self,name):
        (self.unlock_field if name in self.locked_fields else self.lock_field)(name)
        self.mark_dirty()
    def lock_all_filled(self):
        for name,var in self.vars.items():
            if upper(var.get()):self.locked_fields.add(name)
        for name in self.vars:self.apply_lock_visual(name)
        self.mark_dirty(); self.toast("Locked all filled characteristics")
    def unlock_all(self):
        self.locked_fields.clear()
        for name in self.vars:self.apply_lock_visual(name)
        self.mark_dirty(); self.toast("Unlocked all characteristics")
    def apply_lock_visual(self,name):
        box=self.boxes.get(name); btn=self.lock_buttons.get(name); suggest=self.suggest_buttons.get(name)
        if not box or not btn:return
        if name in self.locked_fields:
            box.configure(state="disabled",fg_color="#EAF3EA",border_color="#8FBF9A")
            btn.configure(text="\U0001f512",fg_color="#188A49",hover_color="#136B38")
            if suggest:suggest.configure(state="disabled",fg_color="#B9CDB9")
        else:
            box.configure(state="normal",fg_color="#FFF9DD",border_color="#D7BD60")
            btn.configure(text="\U0001f513",fg_color="#8A8F98",hover_color="#6E727A")
            if suggest:suggest.configure(state="normal",fg_color="#2FA36B")

    # -- local lock persistence (piggybacks on the recovery journal) ---------
    def load_locks(self):
        """Loads per-item locked-field sets from the recovery journal, so
        locks survive closing and reopening the tool."""
        data=self.read_recovery() if self.path else {}
        self.locks={key:set(entry.get("locked",[])) for key,entry in data.items()}

    def toggle_theme(self,value):
        ctk.set_appearance_mode("dark" if value=="Dark" else "light")

    def update_progress_dashboard(self):
        if not self.groups:
            self.progress_bar.set(0); self.progress_label.configure(text="No items loaded yet")
            return
        counts=defaultdict(int)
        for rows in self.groups.values():
            counts[self.group_status(rows)]+=1
        total=len(self.groups); done=counts.get("DONE",0)
        self.progress_bar.set(done/total if total else 0)
        elapsed_hours=max((datetime.now()-self.session_start).total_seconds()/3600,1/3600)
        rate=self.session_done_count/elapsed_hours
        throughput=f"  \u2022  session: {self.session_done_count} done ({rate:.0f}/hr)" if self.session_done_count else ""
        self.progress_label.configure(text=f"{done}/{total} done  \u2022  {counts.get('PENDING',0)} pending  \u2022  "
                                            f"{counts.get('QUERY',0)} query  \u2022  {counts.get('OUT OF SCOPE',0)} out of scope"
                                            f"{throughput}")

    def export_summary_report(self):
        if not self.path:
            messagebox.showinfo(APP,"Load an allocation workbook first.")
            return
        from openpyxl import Workbook
        counts=defaultdict(int)
        for rows in self.groups.values():
            counts[self.group_status(rows)]+=1
        wb=Workbook(); ws=wb.active; ws.title="Summary"
        ws.append(["Brazil Item Coding Studio - Summary Report","Created by Sumit Mondal"])
        ws.append(["Generated",datetime.now().isoformat(timespec="seconds")])
        ws.append(["Allocation file",str(self.path)])
        ws.append([])
        ws.append(["Status","Item count"])
        for status in ["DONE","PENDING","QUERY","OUT OF SCOPE","MIXED"]:
            if counts.get(status):ws.append([status,counts[status]])
        ws.append([])
        ws.append(["Total items",len(self.groups)])
        ws.append(["Total characteristics tracked",len(self.fields)])
        out_path=self.path.parent/f"{self.path.stem}_summary_report.xlsx"
        wb.save(out_path)
        messagebox.showinfo(APP,f"Summary report saved:\n{out_path}")

    # -- navigation / display --------------------------------------------------
    def apply_filter(self,reset=False):
        query=upper(self.search_var.get())
        status_filter=self.filter_menu.get()
        keys=[]
        for key,rows in self.groups.items():
            if status_filter!="ALL" and self.group_status(rows)!=status_filter:continue
            if query and query not in upper(key):
                code=text(self.ws.cell(rows[0],self.hmap.get("External Code",0)).value) if self.hmap.get("External Code") else ""
                if query not in upper(code):continue
            keys.append(key)
        self.filtered=keys
        if reset or self.position>=len(self.filtered):self.position=0
        self.show()

    def group_status(self,rows):
        col=self.hmap.get("Status")
        if not col:return "PENDING"
        values={upper(self.ws.cell(r,col).value) for r in rows}
        values.discard("")
        return next(iter(values)) if len(values)==1 else ("PENDING" if not values else "MIXED")

    def current_key(self):
        return self.filtered[self.position] if self.filtered else None
    def current_rows(self):
        return self.groups.get(self.current_key(),[])

    def show(self):
        if not self.filtered:
            self.status.configure(text="No items match the current filter")
            return
        self.loading=True
        rows=self.current_rows()
        desc=text(self.ws.cell(rows[0],self.hmap[DESCRIPTION_HEADER]).value)
        code=text(self.ws.cell(rows[0],self.hmap.get("External Code",0)).value) if self.hmap.get("External Code") else ""
        self.external_code_lbl.configure(text=f"{code}  ({len(rows)} record(s) in this group)")
        self.original_desc_lbl.configure(text=desc)
        self.refresh_translation()
        status_col=self.hmap.get("Status")
        status_value=upper(self.ws.cell(rows[0],status_col).value) if status_col else ""
        self.item_status_var.set(status_value or "PENDING")
        comment_col=self.hmap.get("Comment")
        self.comment.delete("1.0","end")
        if comment_col:
            self.comment.insert("1.0",text(self.ws.cell(rows[0],comment_col).value))
        self.locked_fields=set(self.locks.get(self.current_key(),set()))
        for f in self.fields:
            source_val=""
            if f.source_col:
                values={text(self.ws.cell(r,f.source_col).value) for r in rows}
                values.discard("")
                source_val=next(iter(values)) if len(values)==1 else (f"{len(values)} values" if values else "")
            self.source_labels[f.name].configure(text=source_val)
            proposed_values={upper(self.ws.cell(r,f.proposed_col).value) for r in rows}
            proposed_values.discard("")
            loaded=next(iter(proposed_values)) if len(proposed_values)==1 else ""
            self.vars[f.name].set(loaded); self.boxes[f.name].last_valid=loaded
            self.apply_lock_visual(f.name)
        self.loading=False; self.dirty=False
        self.status.configure(text=f"Item {self.position+1} of {len(self.filtered)} | {len(self.groups)} total items loaded")

    def mark_dirty(self):
        if not self.loading:self.dirty=True
    def item_status_changed(self,value=None):
        self.mark_dirty()
    def toast(self,msg):
        t=ctk.CTkLabel(self,text=msg,fg_color="#0B4D2C",text_color="white",corner_radius=7); t.place(relx=.5,rely=.94,anchor="center")
        self.after(1600,t.destroy)

    def next(self,pending=False):
        if self.dirty and not messagebox.askyesno(APP,"Discard unsaved changes and move to next item?"):return
        if self.position<len(self.filtered)-1:self.position+=1;self.show()
    def previous(self):
        if self.dirty and not messagebox.askyesno(APP,"Discard unsaved changes and move to previous item?"):return
        if self.position>0:self.position-=1;self.show()

    # -- save / mark done --------------------------------------------------
    def save(self,silent=False):
        if not self.ws or not self.current_rows():
            if not silent:messagebox.showerror(APP,"No allocation item is loaded.")
            return False
        status=upper(self.item_status_var.get())
        comment=self.comment.get("1.0","end-1c").strip()
        if status in {"QUERY","OUT OF SCOPE"} and not comment:
            if not silent:messagebox.showwarning(APP,f"Comment is mandatory before saving an item with Status = {status}.")
            return False
        rows=self.current_rows()
        for f in self.fields:
            value=upper(self.vars[f.name].get())
            for r in rows:
                self.ws.cell(r,f.proposed_col).value=value or None
        status_col=self.hmap.get("Status");date_col=self.hmap.get("Date");comment_col=self.hmap.get("Comment")
        for r in rows:
            if status_col:self.ws.cell(r,status_col).value=None if status=="PENDING" else status
            if date_col and status not in ("","PENDING"):self.ws.cell(r,date_col).value=date.today()
            if comment_col:self.ws.cell(r,comment_col).value=comment or None
        try:
            self.wb.save(self.path)
        except PermissionError:
            if not silent:messagebox.showerror(APP,"Could not save - the workbook is open or locked in Excel. Close it and try again.")
            return False
        self.learn_from_allocation()  # refresh brand/manufacturer knowledge with the newly saved values
        self.save_local_recovery()
        self.dirty=False
        self.update_progress_dashboard()
        if not silent:self.toast("Saved")
        return True

    def mark_done(self):
        self.session_done_count=getattr(self,"session_done_count",0)+1
        self.item_status_var.set("DONE")
        if self.save(silent=False):
            self.toast("Done and saved")
            self.next()

    # -- similar-item detection + bulk apply ---------------------------------
    def find_similar_items(self,threshold=0.55):
        """Other items whose description looks like the same product in a
        different size/color (same brand+model tokens, size stripped)."""
        key=self.current_key()
        if not key:return []
        scored=[]
        for other_key in self.groups:
            if other_key==key:continue
            score=similarity_score(key,other_key)
            if score>=threshold:
                scored.append((score,other_key))
        scored.sort(key=lambda x:-x[0])
        return scored

    def open_similar_items_dialog(self):
        if not self.current_key():
            messagebox.showinfo(APP,"Load an item first.")
            return
        matches=self.find_similar_items()
        win=Toplevel(self); win.title("Similar Items"); win.geometry("640x520"); win.attributes("-topmost",True)
        ctk.CTkLabel(win,text="\U0001f517 Items that look like the same product (size/color variants)",
                    font=ctk.CTkFont(size=14,weight="bold")).pack(anchor="w",padx=14,pady=(12,2))
        ctk.CTkLabel(win,text=f"Current item: {self.current_key()}",font=ctk.CTkFont(size=10),
                    text_color="#666",wraplength=600,justify="left").pack(anchor="w",padx=14,pady=(0,8))
        if not matches:
            ctk.CTkLabel(win,text="No similar items found in this workbook.",font=ctk.CTkFont(size=11)).pack(padx=14,pady=20)
            return
        list_frame=ScrollableFrame(win,fg_color="white"); list_frame.pack(fill="both",expand=True,padx=14)
        check_vars={}
        for score,other_key in matches:
            row=ctk.CTkFrame(list_frame.content,fg_color="#F5F8FC",border_width=1,border_color="#D9E6EF")
            row.pack(fill="x",pady=2,padx=2)
            var=BooleanVar(value=True); check_vars[other_key]=var
            ctk.CTkCheckBox(row,text="",variable=var,width=20).pack(side="left",padx=8,pady=6)
            status=self.group_status(self.groups[other_key])
            ctk.CTkLabel(row,text=f"[{int(score*100)}% match, {status}]  {other_key}",anchor="w",wraplength=520,
                        justify="left",font=ctk.CTkFont(size=10)).pack(side="left",padx=4,pady=6,fill="x",expand=True)
        button_row=ctk.CTkFrame(win,fg_color="transparent"); button_row.pack(fill="x",padx=14,pady=10)
        def do_copy():
            targets=[k for k,v in check_vars.items() if v.get()]
            if not targets:
                messagebox.showinfo(APP,"No items selected.")
                return
            n=self.copy_values_to_items(targets)
            messagebox.showinfo(APP,f"Copied proposed values to {n} item(s). TAMANHO was left for each item's "
                                      "own auto-suggestion since size is exactly what differs between variants.")
            win.destroy(); self.update_progress_dashboard()
        ctk.CTkButton(button_row,text="Copy My Values to Selected",command=do_copy,height=34,
                     fg_color="#2FA36B",hover_color="#238754").pack(side="left")
        ctk.CTkLabel(button_row,text="Locked fields on target items are never overwritten.",
                    font=ctk.CTkFont(size=9),text_color="#888").pack(side="left",padx=10)

    def copy_values_to_items(self,target_keys):
        """Copies every proposed value from the CURRENT item onto each target
        item, except TAMANHO (size is precisely what should differ between
        variants) and any field that is locked on the target. Returns the
        number of items actually updated."""
        source_values={f.name:upper(self.vars[f.name].get()) for f in self.fields if upper(self.vars[f.name].get())}
        source_values.pop(TAMANHO_KEY,None)
        if not source_values:
            return 0
        field_by_name={f.name:f for f in self.fields}
        updated=0
        for key in target_keys:
            rows=self.groups.get(key)
            if not rows:continue
            target_locks=self.locks.get(key,set())
            changed=False
            for name,value in source_values.items():
                if name in target_locks:continue
                f=field_by_name.get(name)
                if not f:continue
                for r in rows:self.ws.cell(r,f.proposed_col).value=value
                changed=True
            if changed:updated+=1
        if updated:
            try:self.wb.save(self.path)
            except PermissionError:
                messagebox.showerror(APP,"Copied in memory, but the workbook is locked in Excel - "
                                          "close it and use Save Draft to write the copies out.")
        self.learn_from_allocation()
        return updated

    # -- quality audit --------------------------------------------------------
    def run_quality_audit(self):
        if not self.groups:
            messagebox.showinfo(APP,"Load an allocation workbook first.")
            return
        field_by_name={f.name:f for f in self.fields}
        guideline_by_name=self.handbook.guidelines
        issues=[]  # (item_key, external_code, field, issue)
        for key,rows in self.groups.items():
            code=text(self.ws.cell(rows[0],self.hmap.get("External Code",0)).value) if self.hmap.get("External Code") else ""
            status=self.group_status(rows)
            values={}
            for f in self.fields:
                v={upper(self.ws.cell(r,f.proposed_col).value) for r in rows}
                v.discard("")
                values[f.name]=next(iter(v)) if len(v)==1 else ("" if not v else "MIXED")
            # 1. Enumerated-value fields must use one of the handbook's listed values, once filled.
            for name,g in guideline_by_name.items():
                val=values.get(name,"")
                if val and g.enum_values and val not in g.enum_values:
                    issues.append((key,code,name,f"'{val}' is not one of the handbook's valid values ({', '.join(g.enum_values)})"))
            # 2. ITM_DESC length check (Learning Priority quality rule: <=60 chars).
            desc_val=values.get(ITEM_DESC_CHAR,"")
            if desc_val and len(desc_val)>60:
                issues.append((key,code,ITEM_DESC_CHAR,f"description is {len(desc_val)} characters - handbook limit is 60"))
            # 3. Fixed-constant fields must match their single required value.
            for name,expected in FIXED_VALUE_FIELDS.items():
                val=values.get(name,"")
                if val and val!=expected:
                    issues.append((key,code,name,f"'{val}' should always be '{expected}' for this LPC"))
            # 4. DONE items should have every non-blank-by-design field filled.
            if status=="DONE":
                for f in self.fields:
                    if f.name in ALWAYS_BLANK_FIELDS:continue
                    if not values.get(f.name,""):
                        issues.append((key,code,f.name,"blank on an item marked DONE"))
            # 5. QUERY / OUT OF SCOPE items must carry a comment (should already be enforced at save time).
            if status in {"QUERY","OUT OF SCOPE"}:
                comment_col=self.hmap.get("Comment")
                comment=text(self.ws.cell(rows[0],comment_col).value) if comment_col else ""
                if not comment:
                    issues.append((key,code,"Comment",f"missing - required for Status={status}"))
        self.show_quality_audit_results(issues)

    def show_quality_audit_results(self,issues):
        win=Toplevel(self); win.title("Quality Audit Results"); win.geometry("760x560"); win.attributes("-topmost",True)
        ctk.CTkLabel(win,text="\U0001f50d Quality Audit",font=ctk.CTkFont(size=16,weight="bold")).pack(anchor="w",padx=14,pady=(12,2))
        summary=f"{len(issues)} issue(s) found across {len(self.groups)} item(s)." if issues else \
                 f"No issues found across {len(self.groups)} item(s). Nice work."
        ctk.CTkLabel(win,text=summary,font=ctk.CTkFont(size=11,weight="bold"),
                    text_color="#B0791F" if issues else "#188A49").pack(anchor="w",padx=14,pady=(0,8))
        if issues:
            list_frame=ScrollableFrame(win,fg_color="white"); list_frame.pack(fill="both",expand=True,padx=14,pady=(0,8))
            for key,code,field,issue in issues[:400]:
                row=ctk.CTkFrame(list_frame.content,fg_color="#FFF7E8",border_width=1,border_color="#EAD9AE")
                row.pack(fill="x",pady=2,padx=2)
                ctk.CTkLabel(row,text=f"{code or key[:40]}  \u2192  {field}",font=ctk.CTkFont(size=10,weight="bold"),
                            anchor="w").pack(anchor="w",padx=8,pady=(4,0))
                ctk.CTkLabel(row,text=issue,font=ctk.CTkFont(size=9),text_color="#7A5A0B",anchor="w",
                            wraplength=680,justify="left").pack(anchor="w",padx=8,pady=(0,4))
            def export():
                from openpyxl import Workbook
                wb=Workbook(); ws=wb.active; ws.title="Quality Audit"
                ws.append(["Item","External Code","Field","Issue"])
                for key,code,field,issue in issues:
                    ws.append([key,code,field,issue])
                out=self.path.parent/f"{self.path.stem}_quality_audit.xlsx"
                wb.save(out)
                messagebox.showinfo(APP,f"Audit exported:\n{out}")
            ctk.CTkButton(win,text="Export Audit to Excel",command=export,height=32,
                         fg_color="#5A6472",hover_color="#454E59").pack(pady=(0,12))

    # -- local recovery journal ---------------------------------------------
    def recovery_dir(self)->Path:
        d=self.path.parent/"BrazilItemCodingRecovery"
        d.mkdir(exist_ok=True)
        return d
    def recovery_path(self)->Path:
        return self.recovery_dir()/f"{self.path.stem}_recovery.json"
    def read_recovery(self)->dict:
        p=self.recovery_path()
        if p.exists():
            try:return json.loads(p.read_text(encoding="utf-8"))
            except Exception:return {}
        return {}
    def save_local_recovery(self):
        """Write an always-up-to-date recovery snapshot of the current item next
        to the allocation file. This MERGES into any existing snapshot for the
        item - a field that already had a value keeps its last known non-blank
        value in the journal even if it is blank in this save, so accidentally
        clearing-then-saving a field can still be recovered. This never blocks
        or fails a Save - if the journal write fails it is silently skipped."""
        if not self.path or not self.current_key():return
        try:
            data=self.read_recovery()
            rows=self.current_rows()
            code=text(self.ws.cell(rows[0],self.hmap.get("External Code",0)).value) if self.hmap.get("External Code") else ""
            existing_values=data.get(self.current_key(),{}).get("values",{})
            new_values={n:upper(v.get()) for n,v in self.vars.items() if upper(v.get())}
            data[self.current_key()]={
                "external_code":code,
                "timestamp":datetime.now().isoformat(timespec="seconds"),
                "status":upper(self.item_status_var.get()),
                "comment":self.comment.get("1.0","end-1c").strip(),
                "values":{**existing_values,**new_values},
                "locked":sorted(self.locked_fields),
            }
            self.recovery_path().write_text(json.dumps(data,indent=2,ensure_ascii=False),encoding="utf-8")
            self.locks[self.current_key()]=set(self.locked_fields)
        except Exception:
            pass

    def recover_local_work(self):
        if not self.path:
            messagebox.showinfo(APP,"Load an allocation workbook first.")
            return
        data=self.read_recovery()
        if not data:
            messagebox.showinfo(APP,"No local recovery journal was found next to this allocation file.")
            return
        if not messagebox.askyesno(APP,f"{len(data)} recovered item snapshot(s) were found in "
                                        f"'{self.recovery_dir().name}'. Reapply any values that are "
                                        "still missing from the workbook now?"):
            return
        restored_items=0; restored_fields=0
        field_by_name={f.name:f for f in self.fields}
        for key,entry in data.items():
            rows=self.groups.get(key)
            if not rows:
                # fall back to matching by external code, in case the description text changed
                code=entry.get("external_code","")
                if code and self.hmap.get("External Code"):
                    for k,rs in self.groups.items():
                        if text(self.ws.cell(rs[0],self.hmap["External Code"]).value)==code:
                            rows=rs;break
            if not rows:continue
            changed=False
            for name,value in entry.get("values",{}).items():
                f=field_by_name.get(name)
                if not f or not value:continue
                current={upper(self.ws.cell(r,f.proposed_col).value) for r in rows}
                if current=={value}:continue          # already correct - nothing to do
                if any(c for c in current):continue   # already holds a different value - never overwrite it
                for r in rows:self.ws.cell(r,f.proposed_col).value=value
                changed=True; restored_fields+=1
            if changed:restored_items+=1
        if restored_items:
            try:
                self.wb.save(self.path)
            except PermissionError:
                messagebox.showerror(APP,"Recovered values were applied in memory, but the workbook is "
                                          "locked in Excel - close it and use Save Draft to write them out.")
                return
        self.learn_from_allocation(); self.show(); self.update_progress_dashboard()
        messagebox.showinfo(APP,f"Recovery complete: {restored_fields} value(s) restored across {restored_items} item(s).")

    # -- concurrency-safe online save ----------------------------------------
    def save_changes_online(self):
        """Save the current item onto the LATEST copy of the workbook on disk,
        instead of the copy that was in memory when the app opened it. This
        protects work when several coders share the same file on a network
        drive: a teammate's edits to other rows are never overwritten."""
        if not self.ws or not self.current_rows():
            messagebox.showerror(APP,"No allocation item is loaded.")
            return
        status=upper(self.item_status_var.get())
        comment=self.comment.get("1.0","end-1c").strip()
        if status in {"QUERY","OUT OF SCOPE"} and not comment:
            messagebox.showwarning(APP,f"Comment is mandatory before saving an item with Status = {status}.")
            return
        key=self.current_key()
        proposed={f.name:upper(self.vars[f.name].get()) for f in self.fields}
        try:
            fresh_wb=load_workbook(self.path)
            fresh_ws=fresh_wb[fresh_wb.sheetnames[0]]
            fresh_fields,fresh_hmap,fresh_groups=self.parse_sheet(fresh_ws)
        except PermissionError:
            messagebox.showerror(APP,"Could not read the shared file - it may be open elsewhere. Try again shortly.")
            return
        except Exception as exc:
            messagebox.showerror(APP,f"Could not read the latest shared file.\n\n{exc}")
            return
        rows=fresh_groups.get(key)
        if not rows:
            messagebox.showwarning(APP,"This item could no longer be found in the latest shared file "
                                        "(it may have been removed or re-described by a teammate). "
                                        "Nothing was overwritten - reload the allocation file to continue.")
            return
        fresh_field_by_name={f.name:f for f in fresh_fields}
        for name,value in proposed.items():
            f=fresh_field_by_name.get(name)
            if not f:continue
            for r in rows:fresh_ws.cell(r,f.proposed_col).value=value or None
        status_col=fresh_hmap.get("Status");date_col=fresh_hmap.get("Date");comment_col=fresh_hmap.get("Comment")
        for r in rows:
            if status_col:fresh_ws.cell(r,status_col).value=None if status=="PENDING" else status
            if date_col and status not in ("","PENDING"):fresh_ws.cell(r,date_col).value=date.today()
            if comment_col:fresh_ws.cell(r,comment_col).value=comment or None
        try:
            fresh_wb.save(self.path)
        except PermissionError:
            messagebox.showerror(APP,"Could not save - the workbook is open or locked in Excel. Close it and try again.")
            return
        # Adopt the freshly-saved copy as the new in-memory workbook so we stay in sync with the shared file.
        self.wb.close()
        self.wb=fresh_wb; self.ws=fresh_ws
        self.fields,self.hmap,self.groups=fresh_fields,fresh_hmap,fresh_groups
        self.learn_from_allocation()
        self.enrich_known_values_from_guidelines()
        self.save_local_recovery()
        keys=list(self.groups.keys())
        self.filtered=[k for k in self.filtered if k in self.groups] or keys
        self.position=self.filtered.index(key) if key in self.filtered else 0
        self.dirty=False; self.build_rows(); self.show()
        self.update_progress_dashboard()
        self.toast("\u2601 Saved online - synced with the latest shared file")

    # -- Ask the Handbook (offline chatbot) ----------------------------------
    # -- AI Assistant settings (stored locally, never bundled/shared) -------
    def get_ai_settings(self):
        s=self.read_settings()
        return s.get("anthropic_api_key",""), s.get("anthropic_model",DEFAULT_AI_MODEL)
    def set_ai_settings(self,api_key,model):
        s=self.read_settings()
        s["anthropic_api_key"]=api_key.strip(); s["anthropic_model"]=(model or DEFAULT_AI_MODEL).strip()
        self.settings_file.write_text(json.dumps(s,indent=2),encoding="utf-8")

    def open_ai_settings_dialog(self,parent):
        api_key,model=self.get_ai_settings()
        win=Toplevel(parent); win.title("AI Assistant Settings"); win.geometry("460x260"); win.attributes("-topmost",True)
        ctk.CTkLabel(win,text="Real-time AI Assistant (optional)",font=ctk.CTkFont(size=14,weight="bold")).pack(anchor="w",padx=14,pady=(14,2))
        ctk.CTkLabel(win,text="Paste your own Anthropic API key to get conversational, real-time answers "
                              "powered by Claude. Without a key, the offline handbook search is used instead - "
                              "it always works, with no setup. Your key is stored only on this computer.",
                    font=ctk.CTkFont(size=10),text_color="#666",wraplength=420,justify="left").pack(anchor="w",padx=14,pady=(0,10))
        ctk.CTkLabel(win,text="API key:",font=ctk.CTkFont(size=10,weight="bold")).pack(anchor="w",padx=14)
        key_var=StringVar(value=api_key)
        ctk.CTkEntry(win,textvariable=key_var,show="\u2022",width=420,height=32).pack(padx=14,pady=(2,10))
        ctk.CTkLabel(win,text="Model (leave default unless your account uses a different one):",
                    font=ctk.CTkFont(size=10,weight="bold")).pack(anchor="w",padx=14)
        model_var=StringVar(value=model)
        ctk.CTkEntry(win,textvariable=model_var,width=420,height=32).pack(padx=14,pady=(2,14))
        def save_and_close():
            self.set_ai_settings(key_var.get(),model_var.get()); win.destroy()
        ctk.CTkButton(win,text="Save",command=save_and_close,width=100,height=32,fg_color="#7A3FD8",hover_color="#622FB0").pack(side="right",padx=14,pady=6)
        ctk.CTkButton(win,text="Clear key (use offline mode)",command=lambda:(key_var.set(""),),width=200,height=32).pack(side="left",padx=14,pady=6)

    def open_chatbot(self):
        win=Toplevel(self); win.title("Ask the Handbook"); win.geometry("620x640"); win.attributes("-topmost",True)
        header=ctk.CTkFrame(win,fg_color="transparent"); header.pack(fill="x",padx=12,pady=(10,0))
        ctk.CTkLabel(header,text="\U0001f4ac Ask the Handbook",font=ctk.CTkFont(size=16,weight="bold"),
                    text_color="#7A3FD8").pack(side="left")
        ctk.CTkButton(header,text="\u2699 AI Settings",command=lambda:self.open_ai_settings_dialog(win),
                     width=120,height=26,fg_color="#5A6472",hover_color="#454E59").pack(side="right")
        mode_label=ctk.CTkLabel(win,text="",font=ctk.CTkFont(size=10,weight="bold")); mode_label.pack(anchor="w",padx=12,pady=(4,2))

        def refresh_mode_label():
            api_key,_=self.get_ai_settings()
            if api_key:
                mode_label.configure(text="\U0001f7e2 Mode: AI Assistant (Claude) - real-time, conversational",text_color="#188A49")
            else:
                mode_label.configure(text="\u26aa Mode: Offline Handbook Search - no setup needed, works without internet",text_color="#666")
        refresh_mode_label()

        transcript_frame=ScrollableFrame(win,fg_color="white"); transcript_frame.pack(fill="both",expand=True,padx=12,pady=(4,0))
        win._photo_refs=[]
        history=[]  # [{"role": "user"/"assistant", "content": "..."}] - only meaningful in AI mode

        def append_bubble(sender,message,images=None):
            is_user=sender=="You"
            bubble=ctk.CTkFrame(transcript_frame.content,fg_color="#EFE6FB" if not is_user else "#E8F3EC",corner_radius=8)
            bubble.pack(fill="x",padx=6,pady=4,anchor="e" if is_user else "w")
            ctk.CTkLabel(bubble,text=sender,font=ctk.CTkFont(size=9,weight="bold"),
                        text_color="#7A3FD8" if not is_user else "#0B7A45").pack(anchor="w",padx=8,pady=(6,0))
            ctk.CTkLabel(bubble,text=message,font=ctk.CTkFont(size=11),justify="left",anchor="w",
                        wraplength=520).pack(anchor="w",padx=8,pady=(0,6))
            for blob in (images or [])[:2]:
                try:
                    img=PhotoImage(data=blob)
                    while img.width()>200 or img.height()>200:img=img.subsample(2,2)
                    win._photo_refs.append(img)
                    tk.Label(bubble,image=img,bd=1,relief="solid").pack(anchor="w",padx=8,pady=(0,6))
                except Exception:
                    continue
            win.update_idletasks()
            transcript_frame.canvas.configure(scrollregion=transcript_frame.canvas.bbox("all"))
            transcript_frame.canvas.yview_moveto(1.0)

        entry_row=ctk.CTkFrame(win,fg_color="transparent"); entry_row.pack(fill="x",padx=12,pady=10)
        question=StringVar()
        entry=ctk.CTkEntry(entry_row,textvariable=question,placeholder_text="Ask a question...",height=34)
        entry.pack(side="left",fill="x",expand=True,padx=(0,6)); entry.focus_set()

        def item_context():
            if self.current_key():
                return self.original_desc_lbl.cget("text")
            return ""

        def images_for_top_hit(q):
            hits=search_handbook(q,self.handbook,top_n=1)
            if not hits:return []
            return self.images_for_field(hits[0][1])

        def ask(event=None):
            q=question.get().strip()
            if not q:return
            append_bubble("You",q); question.set("")
            if not self.handbook.guidelines:
                append_bubble("Assistant","Load the handbook workbook first (Load Handbook) so I have guidelines to search.")
                return
            api_key,model=self.get_ai_settings()
            photos=images_for_top_hit(q)
            if api_key:
                try:
                    reply=call_claude_chatbot(q,self.handbook,api_key,model,history,item_context())
                    history.append({"role":"user","content":q}); history.append({"role":"assistant","content":reply})
                    if len(history)>16:del history[:2]  # keep the conversation window bounded
                    append_bubble("Assistant",reply,photos)
                    return
                except Exception as exc:
                    append_bubble("Assistant",f"(AI Assistant unavailable right now - {exc.__class__.__name__}. "
                                              f"Falling back to offline handbook search.)")
            reply=synthesize_offline_answer(q,self.handbook)
            append_bubble("Assistant",reply,photos)

        entry.bind("<Return>",ask)
        ctk.CTkButton(entry_row,text="Ask",command=ask,width=80,height=34,fg_color="#7A3FD8",hover_color="#622FB0").pack(side="left")
        append_bubble("Assistant","Hi! Ask me anything about the Chinelos coding guidelines - "
                                  "for example: \"what values can COLORAÇÃO have\" or \"when is VERSAO PERSON used\". "
                                  "I can also show you the handbook's reference photos when relevant.")

if __name__=="__main__":
    app=App()
    app.mainloop()
