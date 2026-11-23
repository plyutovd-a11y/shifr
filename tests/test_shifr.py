import shutil
import subprocess
import zipfile
from pathlib import Path

import pytest
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches, Pt

from shifr import crypto
from shifr.document import Document, WrongKeyError, find_all

SECRETS = ["Иванов Иван Иванович", "ООО «Ромашка»", "Договор № 17-42/Б", "Project Zeus", "Северный филиал"]


def make_pptx(path: Path):
    prs = Presentation()
    prs.core_properties.author = "Иванов Иван Иванович"
    prs.core_properties.title = "Project Zeus"
    s = prs.slides.add_slide(prs.slide_layouts[1])
    s.shapes.title.text = "Отчёт: Project Zeus"
    tf = s.placeholders[1].text_frame
    tf.text = "Ответственный: "
    r = tf.paragraphs[0].add_run()
    r.text = "Иванов Иван Иванович"
    r.font.bold = True
    r2 = tf.paragraphs[0].add_run()
    r2.text = ", контрагент ООО «Ромашка»."
    p = tf.add_paragraph()
    p.text = "Договор № 17-42/Б\vподписан 01.02.2026"
    s.notes_slide.notes_text_frame.text = "Не забыть: Иванов Иван Иванович"

    s2 = prs.slides.add_slide(prs.slide_layouts[5])
    s2.shapes.title.text = "Филиалы"
    tbl = s2.shapes.add_table(2, 2, Inches(1), Inches(2), Inches(6), Inches(1)).table
    tbl.cell(0, 0).text = "Филиал"
    tbl.cell(0, 1).text = "Выручка"
    tbl.cell(1, 0).text = "Северный филиал"
    tbl.cell(1, 1).text = "12 500 000"
    grp = s2.shapes.add_group_shape()
    tb = grp.shapes.add_textbox(Inches(1), Inches(4), Inches(4), Inches(1))
    tb.text_frame.text = "В группе: Project Zeus"
    tb.text_frame.paragraphs[0].runs[0].font.size = Pt(28)

    s3 = prs.slides.add_slide(prs.slide_layouts[5])
    cd = CategoryChartData()
    cd.categories = ["Северный филиал", "Южный"]
    cd.add_series("Выручка", (10, 20))
    gf = s3.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1), Inches(6), Inches(4), cd)
    gf.chart.has_title = True
    gf.chart.chart_title.text_frame.text = "Выручка ООО «Ромашка»"
    prs.save(str(path))


def all_text(path: Path) -> str:
    with zipfile.ZipFile(path) as z:
        return "\n".join(z.read(n).decode("utf-8", "replace") for n in z.namelist())


def texts(path: Path, include_notes=False):
    return [it.text for it in Document(path).visible_items(include_notes)]


def select_secrets(doc: Document):
    sel = {}
    for s in SECRETS:
        for k, v in find_all(doc.visible_items(), s).items():
            sel.setdefault(k, []).extend(v)
    return sel


# ---------------------------------------------------------------- crypto
def test_key_roundtrip_and_typo():
    k = crypto.Key.generate()
    s = k.to_string()
    assert s.startswith("SHIFR1-")
    assert crypto.Key.from_string(s.lower().replace("-", " ")) == k
    bad = s[:-3] + ("A" if s[-3] != "A" else "B") + s[-2:]
    with pytest.raises(crypto.KeyFormatError):
        crypto.Key.from_string(bad)
    with pytest.raises(crypto.KeyFormatError):
        crypto.Key.from_string(s[:-4])


def test_fragment_shape_preserved():
    k = crypto.Key.generate()
    plain = "Иванов И.И., тел. 8-900-123-45-67, Zeus"
    enc = crypto.encrypt_fragment(k, 7, plain)
    m = crypto.MARKER_RE.fullmatch(enc)
    body = m.group(2)
    assert len(body) == len(plain) and body != plain
    for a, b in zip(plain, body):
        assert a.isdigit() == b.isdigit() and a.isupper() == b.isupper() and a.isspace() == b.isspace()
        if not a.isalnum():
            assert a == b
    dec = crypto.decrypt_marker(k, m.group(1), body)
    assert dec.ok and dec.text == plain and dec.frag_id == 7
    assert not crypto.decrypt_marker(crypto.Key.generate(), m.group(1), body).ok
    damaged = body[:3] + ("ж" if body[3] != "ж" else "з") + body[4:]
    assert not crypto.decrypt_marker(k, m.group(1), damaged).ok


def test_vault():
    k = crypto.Key.generate()
    blob = crypto.vault_dump(k, {1: {"text": "секрет", "where": "Слайд 1"}})
    assert crypto.vault_load(k, blob)[1]["text"] == "секрет"
    with pytest.raises(ValueError):
        crypto.vault_load(crypto.Key.generate(), blob)


# ---------------------------------------------------------------- pptx
@pytest.fixture
def src(tmp_path):
    p = tmp_path / "deck.pptx"
    make_pptx(p)
    return p


def encrypt(src: Path, tmp_path: Path, **kw):
    doc = Document(src)
    key = crypto.Key.generate()
    rep = doc.encrypt(select_secrets(doc), key, **kw)
    out = tmp_path / "deck.enc.pptx"
    doc.save(out)
    return key, rep, out


def test_no_secret_leaks(src, tmp_path):
    key, rep, out = encrypt(src, tmp_path)
    blob = all_text(out)
    for s in SECRETS:
        assert s not in blob, s
    assert rep.fragments >= 8
    assert any("изображ" in w or "диаграмм" in w for w in rep.warnings)


def test_roundtrip(src, tmp_path):
    before = texts(src)
    key, rep, out = encrypt(src, tmp_path)
    enc_texts = texts(out)
    assert [len(t) for t in enc_texts] != [len(t) for t in before]  # метки добавились
    doc = Document(out)
    drep = doc.decrypt(key)
    assert not drep.damaged and not drep.missing
    dec = tmp_path / "deck.dec.pptx"
    doc.save(dec)
    assert texts(dec) == before


def test_formatting_preserved(src, tmp_path):
    key, _, out = encrypt(src, tmp_path)
    doc = Document(out)
    doc.decrypt(key)
    dec = tmp_path / "d.pptx"
    doc.save(dec)
    runs = Presentation(str(dec)).slides[0].placeholders[1].text_frame.paragraphs[0].runs
    bold = [r for r in runs if r.font.bold]
    assert bold and bold[0].text == "Иванов Иван Иванович"


def test_wrong_key(src, tmp_path):
    _, _, out = encrypt(src, tmp_path)
    with pytest.raises(WrongKeyError):
        Document(out).decrypt(crypto.Key.generate())


def test_damaged_and_missing_with_vault(src, tmp_path):
    key, rep, out = encrypt(src, tmp_path)
    vault = crypto.vault_load(key, crypto.vault_dump(key, rep.vault))
    doc = Document(out)
    # «ИИ» испортил один фрагмент и удалил другой
    hit = [it for it in doc.items if crypto.OPEN in it.text]
    segs = [s for s in hit[0].segments if s is not None]
    h = next(i for i, s in enumerate(segs) if crypto.SEP in (s.text or ""))
    seg = next(s for s in segs[h + 1 :] if any(c.isalpha() for c in (s.text or "")))
    i = next(k for k, c in enumerate(seg.text) if c.isalpha())
    seg.text = seg.text[:i] + ("Щ" if seg.text[i] != "Щ" else "Ж") + seg.text[i + 1 :]
    for s in hit[-1].segments:
        if s is not None:
            s.text = "удалено ИИ"
    tmp = tmp_path / "ai.pptx"
    doc.save(tmp)
    doc = Document(tmp)
    drep = doc.decrypt(key, vault)
    assert len(drep.damaged) == 1 and drep.damaged[0]["from_backup"]
    assert drep.missing and all(m["text"] for m in drep.missing)
    # без резервной копии — испорченный остаётся в отчёте с «примерной» расшифровкой
    drep2 = Document(tmp).decrypt(key)
    assert len(drep2.damaged) == 1 and not drep2.damaged[0]["from_backup"]


@pytest.mark.skipif(not shutil.which("soffice"), reason="LibreOffice не установлен")
def test_libreoffice_resave(src, tmp_path):
    before = texts(src)
    key, _, out = encrypt(src, tmp_path)
    lo = tmp_path / "lo"
    lo.mkdir()
    subprocess.run(
        ["soffice", "--headless", "--convert-to", "pptx", "--outdir", str(lo), str(out)],
        check=True, capture_output=True, timeout=180,
    )
    resaved = lo / out.name
    doc = Document(resaved)
    drep = doc.decrypt(key)
    assert not drep.damaged, drep.damaged
    dec = tmp_path / "lo.dec.pptx"
    doc.save(dec)
    after = set(texts(dec))
    for s in SECRETS:
        assert any(s in t for t in after), s


def test_marker_parts_are_tiny(src, tmp_path):
    key, _, out = encrypt(src, tmp_path)
    prs = Presentation(str(out))
    small, big = 0, 0
    for slide in prs.slides:
        for r in slide.part._element.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}r"):
            txt = r.findtext("{http://schemas.openxmlformats.org/drawingml/2006/main}t") or ""
            if crypto.OPEN in txt or crypto.CLOSE in txt:
                rpr = r.find("{http://schemas.openxmlformats.org/drawingml/2006/main}rPr")
                if rpr is not None and rpr.get("sz") == "100" and len(txt.strip(crypto.OPEN + crypto.CLOSE + crypto.SEP + crypto.B32)) == 0:
                    small += 1
                else:
                    big += 1
    assert small > 0 and big == 0
