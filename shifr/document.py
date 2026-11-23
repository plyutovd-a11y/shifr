"""Работа с .pptx: поиск текста, шифрование выбранных фрагментов, расшифровка, очистка."""
from __future__ import annotations

import copy
import re
from dataclasses import dataclass, field
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.opc.package import XmlPart

from . import crypto

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
A = "{%s}" % NS["a"]
P = "{%s}" % NS["p"]
C = "{%s}" % NS["c"]

KIND_ORDER = {"slide": 0, "chart": 1, "smartart": 2, "notes": 3, "master": 4}
KIND_LABEL = {
    "slide": "",
    "chart": "диаграмма",
    "smartart": "SmartArt",
    "notes": "заметки",
    "master": "образец/макет",
}


@dataclass
class Item:
    """Абзац (или подпись диаграммы), текст которого можно зашифровать."""

    id: str
    slide: int  # 0 — не привязан к слайду (образец/макет)
    kind: str
    shape: str
    segments: list  # элементы с .text (a:t, c:v) или None для разрыва строки
    order: tuple = ()

    @property
    def text(self) -> str:
        return "".join("\n" if s is None else (s.text or "") for s in self.segments)

    @property
    def where(self) -> str:
        parts = []
        if self.slide:
            parts.append(f"Слайд {self.slide}")
        if KIND_LABEL[self.kind]:
            parts.append(KIND_LABEL[self.kind])
        if self.shape:
            parts.append(self.shape)
        return " · ".join(parts)

    def write(self, chars: list[list[str]]):
        for seg, cs in zip(self.segments, chars):
            if seg is not None:
                seg.text = "".join(cs)

    def paragraphs(self):
        out = []
        for seg in self.segments:
            if seg is None:
                continue
            node = seg
            while node is not None and _local(node) != "p":
                node = node.getparent()
            if node is not None and etree.QName(node).namespace == NS["a"] and node not in out:
                out.append(node)
        return out

    def char_map(self):
        """Список символов по сегментам и отображение позиции -> (сегмент, смещение)."""
        chars, pos = [], []
        for si, seg in enumerate(self.segments):
            t = "\n" if seg is None else (seg.text or "")
            chars.append(list(t))
            pos.extend((si, oi) for oi in range(len(t)))
        return chars, pos


@dataclass
class Report:
    warnings: list[str] = field(default_factory=list)
    restored: int = 0
    damaged: list[dict] = field(default_factory=list)  # {id, where, text, from_backup}
    missing: list[dict] = field(default_factory=list)  # {id, where, text}
    fragments: int = 0
    vault: dict = field(default_factory=dict)


# Служебные части метки (⟦номер· и ⟧) пишутся отдельными фрагментами шрифтом 1 pt,
# чтобы почти не занимать места на слайде.
MARKER_SZ = "100"
_MARK_PIECE_RE = re.compile("(%s[%s%s]+[·•∙:|]|%s)" % (crypto.OPEN, crypto.B32, crypto.B32.lower(), crypto.CLOSE))


def _shrink_markers(par):
    for r in list(par.findall(A + "r")):
        t = r.find(A + "t")
        if t is None or not t.text or not _MARK_PIECE_RE.search(t.text):
            continue
        pieces = [x for x in _MARK_PIECE_RE.split(t.text) if x]
        if len(pieces) == 1 and _MARK_PIECE_RE.fullmatch(pieces[0]) and _is_small(r):
            continue
        anchor = r
        for piece in pieces:
            nr = copy.deepcopy(r)
            nr.find(A + "t").text = piece
            if _MARK_PIECE_RE.fullmatch(piece):
                rpr = nr.find(A + "rPr")
                if rpr is None:
                    rpr = etree.SubElement(nr, A + "rPr")
                    nr.remove(rpr)
                    nr.insert(0, rpr)
                rpr.set("sz", MARKER_SZ)
            anchor.addnext(nr)
            anchor = nr
        par.remove(r)


def _is_small(r) -> bool:
    rpr = r.find(A + "rPr")
    return rpr is not None and rpr.get("sz") == MARKER_SZ


def _drop_empty_runs(par):
    runs = par.findall(A + "r")
    for r in runs:
        t = r.find(A + "t")
        if t is not None and not t.text and len(par.findall(A + "r")) > 1:
            par.remove(r)


class WrongKeyError(Exception):
    pass


def _local(el) -> str:
    return etree.QName(el).localname


def _shape_name(el) -> tuple[str, bool]:
    """Имя фигуры, в которой лежит элемент, и признак «это заполнитель»."""
    node = el
    while node is not None:
        if _local(node) in ("sp", "graphicFrame", "cxnSp", "pic"):
            pr = node.find("./*/" + P + "cNvPr")
            if pr is None:
                pr = node.find(".//" + P + "cNvPr")
            is_ph = node.find("./*/" + P + "nvPr/" + P + "ph") is not None
            return (pr.get("name", "") if pr is not None else ""), is_ph
        node = node.getparent()
    return "", False


class Document:
    def __init__(self, path: str | Path):
        self.path = Path(path)
        self.prs = Presentation(str(path))
        self.package = self.prs.part.package
        self._roots: dict[str, tuple] = {}  # partname -> (part, root, is_xml_part)
        self.items: list[Item] = []
        self._scan()

    # ---------------------------------------------------------------- сканирование
    def _root(self, part):
        name = str(part.partname)
        if name not in self._roots:
            if isinstance(part, XmlPart):
                self._roots[name] = (part, part._element, True)
            else:
                self._roots[name] = (part, etree.fromstring(part.blob), False)
        return self._roots[name][1]

    def _slide_parts(self):
        """partname -> (номер слайда, вид) для слайдов и всего, что к ним привязано."""
        owner = {}
        for i, slide in enumerate(self.prs.slides, 1):
            owner[str(slide.part.partname)] = (i, "slide")
            for rel in slide.part.rels.values():
                if rel.is_external:
                    continue
                pn = str(rel.target_part.partname)
                if "/notesSlides/" in pn:
                    owner[pn] = (i, "notes")
                elif "/charts/chart" in pn:
                    owner[pn] = (i, "chart")
                elif "/diagrams/data" in pn or "/diagrams/drawing" in pn:
                    owner[pn] = (i, "smartart")
        return owner

    def _scan(self):
        self.items.clear()
        owner = self._slide_parts()
        for part in self.package.iter_parts():
            pn = str(part.partname)
            if pn in owner:
                slide, kind = owner[pn]
            elif "/slideMasters/" in pn or "/slideLayouts/" in pn:
                slide, kind = 0, "master"
            elif "/charts/chart" in pn:
                slide, kind = 0, "chart"
            elif "/diagrams/data" in pn or "/diagrams/drawing" in pn:
                slide, kind = 0, "smartart"
            else:
                continue
            root = self._root(part)
            for idx, p in enumerate(root.iter(A + "p")):
                shape, is_ph = _shape_name(p)
                if kind == "master" and is_ph:
                    continue  # подсказки заполнителей макетов не трогаем
                segs = []
                for ch in p:
                    ln = _local(ch)
                    if ln in ("r", "fld"):
                        t = ch.find(A + "t")
                        if t is not None:
                            segs.append(t)
                    elif ln == "br":
                        segs.append(None)
                if kind == "notes" and shape == "" and not segs:
                    continue
                self.items.append(
                    Item(f"{pn}#p{idx}", slide, kind, shape, segs, (slide or 10**6, KIND_ORDER[kind], pn, idx))
                )
            if kind == "chart":
                for idx, v in enumerate(root.iter(C + "v")):
                    par = v.getparent()
                    if par is None or _local(par) != "pt":
                        continue
                    cache = par.getparent()
                    if cache is None or _local(cache) != "strCache":
                        continue
                    self.items.append(
                        Item(f"{pn}#v{idx}", slide, kind, "подпись данных", [v], (slide or 10**6, 1, pn, 10**6 + idx))
                    )
        self.items.sort(key=lambda it: it.order)

    def visible_items(self, include_notes: bool = True) -> list[Item]:
        return [it for it in self.items if it.text.strip() and (include_notes or it.kind != "notes")]

    def item(self, item_id: str) -> Item:
        for it in self.items:
            if it.id == item_id:
                return it
        raise KeyError(item_id)

    # ---------------------------------------------------------------- шифрование
    @staticmethod
    def _normalize(text: str, ranges) -> list[tuple[int, int]]:
        cleaned = []
        for s, e in sorted(ranges):
            s, e = max(0, s), min(len(text), e)
            while s < e and text[s].isspace():
                s += 1
            while e > s and text[e - 1].isspace():
                e -= 1
            if s >= e:
                continue
            if cleaned and s <= cleaned[-1][1]:
                cleaned[-1] = (cleaned[-1][0], max(cleaned[-1][1], e))
            else:
                cleaned.append((s, e))
        # разрыв строки внутри фрагмента допустим, но фрагмент не должен целиком быть пробелами
        return cleaned

    def encrypt(
        self,
        selections: dict[str, list[tuple[int, int]]],
        key: crypto.Key,
        remove_notes: bool = True,
        remove_alt_text: bool = True,
    ) -> Report:
        rep = Report()
        frag_id = 0
        plan = []
        for it in self.items:
            if it.id not in selections:
                continue
            if remove_notes and it.kind == "notes":
                continue
            text = it.text
            for s, e in self._normalize(text, selections[it.id]):
                for ch in crypto.FORBIDDEN_CHARS:
                    if ch in text[s:e]:
                        raise ValueError(f"{it.where}: текст содержит служебный символ «{ch}».")
                frag_id += 1
                plan.append((it, s, e, frag_id))

        by_item: dict[str, list] = {}
        for it, s, e, fid in plan:
            by_item.setdefault(it.id, []).append((it, s, e, fid))
        for frags in by_item.values():
            it = frags[0][0]
            text = it.text
            chars, pos = it.char_map()
            for _, s, e, fid in frags:
                head, ct, tail = crypto.marker_parts(key, fid, text[s:e])
                for k in range(s, e):
                    si, oi = pos[k]
                    if it.segments[si] is not None:
                        chars[si][oi] = ct[k - s]
                si, oi = pos[s]
                chars[si][oi] = head + chars[si][oi]
                si, oi = pos[e - 1]
                chars[si][oi] = chars[si][oi] + tail
                rep.vault[fid] = {"text": text[s:e], "where": it.where}
            it.write(chars)
            for par in it.paragraphs():
                _shrink_markers(par)
        rep.fragments = frag_id

        self._sanitize(rep, remove_notes, remove_alt_text)
        self._warn(rep)
        return rep

    # ---------------------------------------------------------------- очистка
    def _sanitize(self, rep: Report, remove_notes: bool, remove_alt_text: bool):
        # комментарии
        dropped_comments = 0
        for part in list(self.package.iter_parts()):
            for rId, rel in list(part.rels.items()):
                rt = rel.reltype.lower()
                if rt.endswith("/comments") or "comment" in rt.rsplit("/", 1)[-1] or rt.endswith("/authors"):
                    part.rels.pop(rId)
                    dropped_comments += 1
        if dropped_comments:
            rep.warnings.append("Комментарии удалены.")

        # заметки
        if remove_notes:
            n = 0
            for slide in self.prs.slides:
                for rId, rel in list(slide.part.rels.items()):
                    if rel.reltype.endswith("/notesSlide"):
                        slide.part.rels.pop(rId)
                        n += 1
            if n:
                rep.warnings.append(f"Заметки докладчика удалены ({n} сл.).")

        # миниатюра, расширенные и пользовательские свойства документа
        for rId, rel in list(self.package._rels.items()):
            if rel.reltype.endswith(("/thumbnail", "/extended-properties", "/custom-properties")):
                self.package._rels.pop(rId)
        cp = self.prs.core_properties
        for attr in ("author", "title", "subject", "keywords", "comments", "last_modified_by",
                     "category", "content_status", "identifier", "version", "language"):
            try:
                setattr(cp, attr, "")
            except Exception:
                pass
        cp.revision = 1

        # замещающий текст (alt) у объектов на слайдах
        if remove_alt_text:
            n = 0
            for slide in self.prs.slides:
                for pr in slide.part._element.iter(P + "cNvPr"):
                    for attr in ("descr", "title"):
                        if pr.get(attr):
                            del pr.attrib[attr]
                            n += 1
            if n:
                rep.warnings.append(f"Замещающий текст (alt) удалён у {n} объектов.")

        # встроенные таблицы данных диаграмм содержат исходные подписи
        n = 0
        for part in list(self.package.iter_parts()):
            if "/charts/chart" not in str(part.partname):
                continue
            for rId, rel in list(part.rels.items()):
                if rel.reltype.endswith("/package") or rel.reltype.endswith("/oleObject"):
                    root = self._root(part)
                    for ext in root.iter(C + "externalData"):
                        ext.getparent().remove(ext)
                    part.rels.pop(rId)
                    n += 1
        if n:
            rep.warnings.append(
                f"У {n} диаграмм удалены встроенные таблицы данных (там остаётся исходный текст). "
                "Данные этих диаграмм больше нельзя редактировать в Excel."
            )

    def _warn(self, rep: Report):
        for i, slide in enumerate(self.prs.slides, 1):
            el = slide.part._element
            pics = len(el.findall(".//" + P + "pic"))
            if pics:
                rep.warnings.append(f"Слайд {i}: изображений — {pics}. Текст на картинках не шифруется, проверьте вручную.")
            if el.findall(".//" + P + "oleObj"):
                rep.warnings.append(f"Слайд {i}: есть внедрённые объекты (OLE) — их содержимое не шифруется.")
            ext = [r for r in slide.part.rels.values() if r.is_external and r.reltype.endswith("/hyperlink")]
            if ext:
                rep.warnings.append(f"Слайд {i}: есть гиперссылки ({len(ext)}) — адреса ссылок не шифруются.")
            kinds = {str(r.target_part.partname) for r in slide.part.rels.values() if not r.is_external}
            if any("/diagrams/" in k for k in kinds):
                rep.warnings.append(
                    f"Слайд {i}: SmartArt хранит текст дважды (данные и отрисовка) — убедитесь, что отмечены обе копии."
                )
            if any("/charts/" in k for k in kinds):
                rep.warnings.append(f"Слайд {i}: у диаграмм числовые значения не шифруются.")

    # ---------------------------------------------------------------- расшифровка
    def decrypt(self, key: crypto.Key, vault: dict[int, dict] | None = None) -> Report:
        rep = Report()
        found_ok, found_bad, found_ids = 0, 0, set()
        pending = []  # (item, chars, pos, match, decrypted)
        for it in self.items:
            text = it.text
            if crypto.OPEN not in text:
                continue
            chars, pos = it.char_map()
            changed = False
            for m in crypto.MARKER_RE.finditer(text):
                try:
                    dec = crypto.decrypt_marker(key, m.group(1), m.group(2))
                except ValueError:
                    continue
                found_ids.add(dec.frag_id)
                if dec.ok:
                    found_ok += 1
                    new = dec.text
                else:
                    found_bad += 1
                    backup = (vault or {}).get(dec.frag_id)
                    rep.damaged.append({
                        "id": dec.frag_id, "where": it.where,
                        "text": backup["text"] if backup else dec.text,
                        "from_backup": bool(backup),
                    })
                    if not backup:
                        continue
                    new = None
                body_s, body_e = m.start(2), m.end(2)
                for k in range(m.start(), m.end()):
                    si, oi = pos[k]
                    if it.segments[si] is None:
                        continue  # разрыв строки оставляем как есть
                    if body_s <= k < body_e:
                        chars[si][oi] = new[k - body_s] if new is not None else ""
                    else:
                        chars[si][oi] = ""
                if new is None:  # восстановление из резервной копии — целиком в начало
                    si, oi = pos[body_s] if body_s < body_e else pos[m.start()]
                    chars[si][oi] = backup["text"]
                changed = True
                rep.restored += 1
            if changed:
                pending.append((it, chars))

        if found_ids and found_ok == 0 and found_bad > 0:
            raise WrongKeyError("Ключ не подходит к этой презентации.")
        if not found_ids:
            raise WrongKeyError("В презентации не найдено зашифрованных фрагментов.")

        for it, chars in pending:
            it.write(chars)
            for par in it.paragraphs():
                _drop_empty_runs(par)

        if vault:
            for fid in sorted(set(vault) - found_ids):
                rep.missing.append({"id": fid, "where": vault[fid].get("where", ""), "text": vault[fid]["text"]})
        else:
            for fid in range(1, max(found_ids)):
                if fid not in found_ids:
                    rep.missing.append({"id": fid, "where": "", "text": ""})
        rep.fragments = len(found_ids)
        return rep

    # ---------------------------------------------------------------- сохранение
    def save(self, path: str | Path):
        for part, root, is_xml in self._roots.values():
            if not is_xml:
                part.blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        self.prs.save(str(path))


def find_all(items: list[Item], needle: str, ignore_case: bool = True) -> dict[str, list[tuple[int, int]]]:
    """Все вхождения строки во всех абзацах."""
    out: dict[str, list[tuple[int, int]]] = {}
    needle = needle.strip()
    if not needle:
        return out
    flags = re.IGNORECASE if ignore_case else 0
    rx = re.compile(re.escape(needle), flags)
    for it in items:
        for m in rx.finditer(it.text):
            out.setdefault(it.id, []).append((m.start(), m.end()))
    return out


def output_path(src: Path, suffix: str) -> Path:
    stem = src.stem
    for s in (".enc", ".dec"):
        if stem.endswith(s):
            stem = stem[: -len(s)]
    cand = src.with_name(f"{stem}{suffix}.pptx")
    n = 2
    while cand.exists():
        cand = src.with_name(f"{stem}{suffix} ({n}).pptx")
        n += 1
    return cand
