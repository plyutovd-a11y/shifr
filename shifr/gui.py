"""Графический интерфейс Shifr (PySide6)."""
from __future__ import annotations

import os
import sys
import traceback
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QFont, QGuiApplication, QTextCharFormat, QTextCursor
from PySide6.QtWidgets import (
    QAbstractItemView, QApplication, QCheckBox, QDialog, QFileDialog, QHBoxLayout, QHeaderView,
    QLabel, QLineEdit, QMainWindow, QMessageBox, QPlainTextEdit, QPushButton, QSplitter,
    QTableWidget, QTableWidgetItem, QTabWidget, QTextEdit, QTreeWidget, QTreeWidgetItem,
    QVBoxLayout, QWidget,
)

from . import __version__, crypto
from .document import Document, WrongKeyError, find_all, output_path

ID_ROLE = Qt.UserRole + 1
MARK_BG = QColor("#ffe08a")


def error_box(parent, title: str, exc: BaseException):
    msg = str(exc) or exc.__class__.__name__
    box = QMessageBox(QMessageBox.Critical, title, msg, parent=parent)
    box.setDetailedText("".join(traceback.format_exception(exc)))
    box.exec()


def mono_font(size: int = 11) -> QFont:
    f = QFont("Monospace")
    f.setStyleHint(QFont.TypeWriter)
    f.setPointSize(size)
    return f


# ============================================================== диалог ключа
class KeyDialog(QDialog):
    def __init__(self, key_str: str, enc_path: Path, warnings: list[str], parent=None):
        super().__init__(parent)
        self.setWindowTitle("Ключ расшифровки — сохраните его")
        self.setModal(True)
        self.resize(720, 420)
        self._enc_path = enc_path
        lay = QVBoxLayout(self)

        lay.addWidget(QLabel(f"Зашифрованная презентация сохранена:\n{enc_path}"))
        warn = QLabel(
            "<b>Сохраните ключ.</b> Программа его нигде не хранит. "
            "Без ключа расшифровать презентацию невозможно — восстановить его нельзя.<br>"
            "Не храните ключ рядом с презентацией и не отправляйте его вместе с ней."
        )
        warn.setWordWrap(True)
        lay.addWidget(warn)

        self.key_edit = QLineEdit(key_str)
        self.key_edit.setReadOnly(True)
        self.key_edit.setFont(mono_font(13))
        lay.addWidget(self.key_edit)

        row = QHBoxLayout()
        b_copy = QPushButton("Скопировать ключ")
        b_copy.clicked.connect(self._copy)
        b_save = QPushButton("Сохранить ключ в файл…")
        b_save.clicked.connect(self._save)
        row.addWidget(b_copy)
        row.addWidget(b_save)
        row.addStretch()
        lay.addLayout(row)
        self.status = QLabel("")
        lay.addWidget(self.status)

        if warnings:
            lay.addWidget(QLabel("Обратите внимание:"))
            w = QPlainTextEdit("\n".join("• " + x for x in warnings))
            w.setReadOnly(True)
            lay.addWidget(w, 1)
        else:
            lay.addStretch(1)

        self.confirm = QCheckBox("Я сохранил ключ в надёжном месте")
        lay.addWidget(self.confirm)
        self.done_btn = QPushButton("Готово")
        self.done_btn.setEnabled(False)
        self.confirm.toggled.connect(self.done_btn.setEnabled)
        self.done_btn.clicked.connect(self.accept)
        lay.addWidget(self.done_btn, alignment=Qt.AlignRight)

    def _copy(self):
        QGuiApplication.clipboard().setText(self.key_edit.text())
        self.status.setText("Ключ скопирован в буфер обмена.")

    def _save(self):
        default = str(Path.home() / (self._enc_path.stem.removesuffix(".enc") + ".ключ.txt"))
        path, _ = QFileDialog.getSaveFileName(self, "Сохранить ключ", default, "Текст (*.txt)")
        if not path:
            return
        p = Path(path)
        if p.parent.resolve() == self._enc_path.parent.resolve():
            if QMessageBox.question(
                self, "Ключ рядом с презентацией",
                "Вы сохраняете ключ в ту же папку, что и зашифрованную презентацию. "
                "Если их отправят вместе, шифрование потеряет смысл. Всё равно сохранить?",
            ) != QMessageBox.Yes:
                return
        p.write_text(self.key_edit.text() + "\n", encoding="utf-8")
        self.status.setText(f"Ключ сохранён: {p}")

    def reject(self):  # Esc и крестик — только после подтверждения
        if self.confirm.isChecked():
            super().reject()
        else:
            self.status.setText("Сначала сохраните ключ и отметьте галочку ниже.")


# ============================================================== шифрование
class EncryptTab(QWidget):
    def __init__(self):
        super().__init__()
        self.path: Path | None = None
        self.doc: Document | None = None
        self.sel: dict[str, list[tuple[int, int]]] = {}
        self._tree_items: dict[str, QTreeWidgetItem] = {}
        self._updating = False

        lay = QVBoxLayout(self)
        top = QHBoxLayout()
        b_open = QPushButton("Открыть презентацию…")
        b_open.clicked.connect(self.choose_file)
        top.addWidget(b_open)
        self.file_label = QLabel("Файл не выбран. Можно перетащить .pptx в окно.")
        top.addWidget(self.file_label, 1)
        lay.addLayout(top)

        split = QSplitter(Qt.Horizontal)
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Текст (галочка — зашифровать абзац целиком)"])
        self.tree.setSelectionMode(QAbstractItemView.SingleSelection)
        self.tree.itemChanged.connect(self._on_check)
        self.tree.currentItemChanged.connect(self._on_current)
        left = QWidget()
        ll = QVBoxLayout(left)
        ll.setContentsMargins(0, 0, 0, 0)
        ll.addWidget(self.tree, 1)
        lr = QHBoxLayout()
        b_all_on = QPushButton("Отметить всё")
        b_all_on.setToolTip("Зашифровать весь текст презентации (каждый абзац целиком)")
        b_all_on.clicked.connect(self.mark_all)
        b_all_off = QPushButton("Снять всё")
        b_all_off.setToolTip("Снять все отметки")
        b_all_off.clicked.connect(self.clear_all)
        lr.addWidget(b_all_on)
        lr.addWidget(b_all_off)
        lr.addStretch()
        ll.addLayout(lr)
        split.addWidget(left)

        right = QWidget()
        rl = QVBoxLayout(right)
        rl.setContentsMargins(0, 0, 0, 0)
        self.where = QLabel("")
        rl.addWidget(self.where)
        self.view = QTextEdit()
        self.view.setReadOnly(True)
        self.view.setPlaceholderText("Выберите абзац слева, выделите мышью часть текста и нажмите кнопку ниже.")
        rl.addWidget(self.view, 1)
        r1 = QHBoxLayout()
        b_mark = QPushButton("Зашифровать выделенное")
        b_mark.clicked.connect(self.mark_selection)
        b_all = QPushButton("Зашифровать все вхождения выделенного")
        b_all.clicked.connect(self.mark_selection_everywhere)
        b_clear = QPushButton("Снять отметки в абзаце")
        b_clear.clicked.connect(self.clear_current)
        r1.addWidget(b_mark)
        r1.addWidget(b_all)
        r1.addWidget(b_clear)
        rl.addLayout(r1)
        r2 = QHBoxLayout()
        r2.addWidget(QLabel("Найти и зашифровать везде:"))
        self.find_edit = QLineEdit()
        self.find_edit.setPlaceholderText("например, ООО «Ромашка» (регистр не важен)")
        self.find_edit.returnPressed.connect(self.mark_found)
        r2.addWidget(self.find_edit, 1)
        b_find = QPushButton("Отметить")
        b_find.clicked.connect(self.mark_found)
        r2.addWidget(b_find)
        rl.addLayout(r2)
        split.addWidget(right)
        split.setSizes([520, 520])
        lay.addWidget(split, 1)

        opts = QHBoxLayout()
        self.cb_notes = QCheckBox("Удалить заметки докладчика")
        self.cb_notes.setChecked(True)
        self.cb_notes.toggled.connect(self._rebuild_tree)
        self.cb_alt = QCheckBox("Удалить замещающий текст (alt)")
        self.cb_alt.setChecked(True)
        self.cb_backup = QCheckBox("Сохранить резервную копию (.shifr)")
        self.cb_backup.setToolTip(
            "Зашифрованный файл с исходными фрагментами. Поможет, если ИИ испортит или удалит часть меток."
        )
        for cb in (self.cb_notes, self.cb_alt, self.cb_backup):
            opts.addWidget(cb)
        opts.addStretch()
        lay.addLayout(opts)

        bottom = QHBoxLayout()
        self.counter = QLabel("Отмечено фрагментов: 0")
        bottom.addWidget(self.counter, 1)
        b_prev = QPushButton("Предпросмотр")
        b_prev.clicked.connect(self.preview)
        self.b_go = QPushButton("Зашифровать и сохранить…")
        self.b_go.clicked.connect(self.run)
        bottom.addWidget(b_prev)
        bottom.addWidget(self.b_go)
        lay.addLayout(bottom)

    # ---------------------------------------------------------- загрузка
    def choose_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Открыть презентацию", "", "PowerPoint (*.pptx)")
        if path:
            self.load(Path(path))

    def load(self, path: Path):
        try:
            doc = Document(path)
        except Exception as e:
            error_box(self, "Не удалось открыть файл", e)
            return
        self.path, self.doc, self.sel = path, doc, {}
        self.file_label.setText(str(path))
        self._rebuild_tree()

    def _rebuild_tree(self):
        if not self.doc:
            return
        self._updating = True
        self.tree.clear()
        self._tree_items.clear()
        groups: dict[str, QTreeWidgetItem] = {}
        for it in self.doc.visible_items(include_notes=not self.cb_notes.isChecked()):
            g = f"Слайд {it.slide}" if it.slide else "Образец слайдов и макеты"
            if g not in groups:
                groups[g] = QTreeWidgetItem(self.tree, [g])
                groups[g].setFlags(Qt.ItemIsEnabled)
                groups[g].setExpanded(True)
            label = it.where.split(" · ", 1)[1] if it.slide and " · " in it.where else it.where
            node = QTreeWidgetItem(groups[g])
            node.setData(0, ID_ROLE, it.id)
            node.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable)
            node.setToolTip(0, f"{it.where}\n\n{it.text}")
            self._tree_items[it.id] = node
            self._paint_node(node, it, label)
        self._updating = False
        self._update_counter()

    def _paint_node(self, node, it, label=None):
        ranges = self.sel.get(it.id, [])
        full = bool(ranges) and Document._normalize(it.text, ranges) == Document._normalize(it.text, [(0, len(it.text))])
        if label is None:
            label = node.data(0, Qt.UserRole + 2) or ""
        node.setData(0, Qt.UserRole + 2, label)
        flat = it.text.replace("\n", " ⏎ ")
        prefix = "" if full or not ranges else "◐ "
        node.setText(0, f"{prefix}{flat}" + (f"   [{label}]" if label else ""))
        node.setCheckState(0, Qt.Checked if full else Qt.Unchecked)
        node.setBackground(0, MARK_BG if ranges else QColor(0, 0, 0, 0))

    # ---------------------------------------------------------- отметки
    def _current_item(self):
        node = self.tree.currentItem()
        if node is None or node.data(0, ID_ROLE) is None or not self.doc:
            return None
        return self.doc.item(node.data(0, ID_ROLE))

    def _on_current(self, *_):
        it = self._current_item()
        if it is None:
            self.view.clear()
            self.where.setText("")
            return
        self.where.setText(it.where)
        self.view.setPlainText(it.text)
        self._highlight(it)

    def _highlight(self, it):
        sels = []
        for s, e in Document._normalize(it.text, self.sel.get(it.id, [])):
            c = QTextCursor(self.view.document())
            c.setPosition(s)
            c.setPosition(e, QTextCursor.KeepAnchor)
            es = QTextEdit.ExtraSelection()
            fmt = QTextCharFormat()
            fmt.setBackground(MARK_BG)
            es.format, es.cursor = fmt, c
            sels.append(es)
        self.view.setExtraSelections(sels)

    def _on_check(self, node, _col):
        if self._updating or node.data(0, ID_ROLE) is None:
            return
        it = self.doc.item(node.data(0, ID_ROLE))
        if node.checkState(0) == Qt.Checked:
            self.sel[it.id] = [(0, len(it.text))]
        else:
            self.sel.pop(it.id, None)
        self._refresh(it.id)

    def _refresh(self, item_id):
        self._updating = True
        node = self._tree_items.get(item_id)
        if node is not None:
            self._paint_node(node, self.doc.item(item_id))
        self._updating = False
        cur = self._current_item()
        if cur is not None and cur.id == item_id:
            self._highlight(cur)
        self._update_counter()

    def _add(self, found: dict):
        for k, ranges in found.items():
            self.sel.setdefault(k, []).extend(ranges)
            self.sel[k] = Document._normalize(self.doc.item(k).text, self.sel[k])
            self._refresh(k)

    def _selected_text(self):
        c = self.view.textCursor()
        return c.selectionStart(), c.selectionEnd()

    def mark_selection(self):
        it = self._current_item()
        if it is None:
            return
        s, e = self._selected_text()
        if s == e:
            QMessageBox.information(self, "Shifr", "Выделите мышью часть текста справа.")
            return
        self._add({it.id: [(s, e)]})

    def mark_selection_everywhere(self):
        it = self._current_item()
        if it is None:
            return
        s, e = self._selected_text()
        if s == e:
            QMessageBox.information(self, "Shifr", "Выделите мышью часть текста справа.")
            return
        self._mark_text(it.text[s:e])

    def mark_found(self):
        self._mark_text(self.find_edit.text())

    def _mark_text(self, needle: str):
        if not self.doc or not needle.strip():
            return
        found = find_all(self.doc.visible_items(not self.cb_notes.isChecked()), needle)
        self._add(found)
        n = sum(len(v) for v in found.values())
        self.counter.setText(self.counter.text() + f"   (найдено «{needle.strip()}»: {n})")

    def mark_all(self):
        if not self.doc:
            return
        for it in self.doc.visible_items(not self.cb_notes.isChecked()):
            self.sel[it.id] = [(0, len(it.text))]
        self._rebuild_tree()
        self._on_current()

    def clear_all(self):
        if not self.doc:
            return
        self.sel.clear()
        self._rebuild_tree()
        self._on_current()

    def clear_current(self):
        it = self._current_item()
        if it is not None:
            self.sel.pop(it.id, None)
            self._refresh(it.id)

    def _active_selection(self):
        visible = {it.id for it in self.doc.visible_items(not self.cb_notes.isChecked())}
        return {k: v for k, v in self.sel.items() if k in visible and v}

    def _update_counter(self):
        if not self.doc:
            return
        n = sum(len(Document._normalize(self.doc.item(k).text, v)) for k, v in self._active_selection().items())
        self.counter.setText(f"Отмечено фрагментов: {n}")

    # ---------------------------------------------------------- действия
    def _encrypted_doc(self, key):
        doc = Document(self.path)  # свежая копия, чтобы не портить исходное состояние
        rep = doc.encrypt(self._active_selection(), key, self.cb_notes.isChecked(), self.cb_alt.isChecked())
        return doc, rep

    def preview(self):
        if not self.doc:
            return
        try:
            doc, _ = self._encrypted_doc(crypto.Key.generate())
        except Exception as e:
            error_box(self, "Ошибка", e)
            return
        lines = [f"[{it.where}]\n{it.text}\n" for it in doc.visible_items(not self.cb_notes.isChecked())]
        dlg = QDialog(self)
        dlg.setWindowTitle("Предпросмотр: так текст увидит ИИ")
        dlg.resize(800, 600)
        l = QVBoxLayout(dlg)
        t = QPlainTextEdit("\n".join(lines))
        t.setReadOnly(True)
        l.addWidget(t)
        b = QPushButton("Закрыть")
        b.clicked.connect(dlg.accept)
        l.addWidget(b, alignment=Qt.AlignRight)
        dlg.exec()

    def run(self):
        if not self.doc:
            QMessageBox.information(self, "Shifr", "Сначала откройте презентацию.")
            return
        if not self._active_selection():
            QMessageBox.information(self, "Shifr", "Не отмечено ни одного фрагмента.")
            return
        default = str(output_path(self.path, ".enc"))
        out, _ = QFileDialog.getSaveFileName(self, "Сохранить зашифрованную презентацию", default, "PowerPoint (*.pptx)")
        if not out:
            return
        out = Path(out)
        if out.resolve() == self.path.resolve():
            QMessageBox.warning(self, "Shifr", "Нельзя перезаписывать исходный файл. Выберите другое имя.")
            return
        key = crypto.Key.generate()
        try:
            doc, rep = self._encrypted_doc(key)
            doc.save(out)
            if self.cb_backup.isChecked():
                backup = out.with_suffix(".shifr")
                backup.write_bytes(crypto.vault_dump(key, rep.vault))
                rep.warnings.insert(0, f"Резервная копия: {backup}")
        except Exception as e:
            error_box(self, "Не удалось зашифровать", e)
            return
        rep.warnings.insert(0, f"Зашифровано фрагментов: {rep.fragments}.")
        KeyDialog(key.to_string(), out, rep.warnings, self).exec()
        del key


# ============================================================== расшифровка
class DecryptTab(QWidget):
    def __init__(self):
        super().__init__()
        self.path: Path | None = None
        lay = QVBoxLayout(self)

        r = QHBoxLayout()
        b = QPushButton("Открыть зашифрованную презентацию…")
        b.clicked.connect(self.choose_file)
        r.addWidget(b)
        self.file_label = QLabel("Файл не выбран. Можно перетащить .pptx в окно.")
        r.addWidget(self.file_label, 1)
        lay.addLayout(r)

        lay.addWidget(QLabel("Ключ:"))
        self.key_edit = QLineEdit()
        self.key_edit.setFont(mono_font(12))
        self.key_edit.setPlaceholderText("SHIFR1-XXXXXX-XXXXXX-…")
        self.key_edit.textChanged.connect(self._check_key)
        lay.addWidget(self.key_edit)
        self.key_status = QLabel("")
        lay.addWidget(self.key_status)

        r2 = QHBoxLayout()
        r2.addWidget(QLabel("Резервная копия (необязательно):"))
        self.backup_edit = QLineEdit()
        self.backup_edit.setPlaceholderText("файл .shifr, если сохраняли")
        r2.addWidget(self.backup_edit, 1)
        b2 = QPushButton("Выбрать…")
        b2.clicked.connect(self.choose_backup)
        r2.addWidget(b2)
        lay.addLayout(r2)

        self.b_go = QPushButton("Расшифровать и сохранить…")
        self.b_go.clicked.connect(self.run)
        lay.addWidget(self.b_go, alignment=Qt.AlignRight)

        self.summary = QLabel("")
        self.summary.setWordWrap(True)
        lay.addWidget(self.summary)
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["№", "Где", "Состояние", "Текст", ""])
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setVisible(False)
        lay.addWidget(self.table, 1)
        lay.addStretch()

    def choose_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Открыть презентацию", "", "PowerPoint (*.pptx)")
        if path:
            self.load(Path(path))

    def load(self, path: Path):
        self.path = path
        self.file_label.setText(str(path))
        if not self.backup_edit.text():
            stem = path.stem
            for s in (".enc", ".dec"):
                stem = stem.removesuffix(s)
            for cand in (path.with_suffix(".shifr"), path.with_name(stem + ".enc.shifr")):
                if cand.exists():
                    self.backup_edit.setText(str(cand))
                    break

    def choose_backup(self):
        path, _ = QFileDialog.getOpenFileName(self, "Резервная копия", "", "Shifr (*.shifr)")
        if path:
            self.backup_edit.setText(path)

    def _check_key(self):
        t = self.key_edit.text().strip()
        if not t:
            self.key_status.setText("")
            return None
        try:
            k = crypto.Key.from_string(t)
        except crypto.KeyFormatError as e:
            self.key_status.setText(f"<span style='color:#b00020'>✗ {e}</span>")
            return None
        self.key_status.setText("<span style='color:#1b7f3b'>✓ Ключ введён корректно</span>")
        return k

    def run(self):
        if not self.path:
            QMessageBox.information(self, "Shifr", "Сначала откройте зашифрованную презентацию.")
            return
        key = self._check_key()
        if key is None:
            QMessageBox.information(self, "Shifr", "Введите корректный ключ.")
            return
        try:
            vault = None
            if self.backup_edit.text().strip():
                vault = crypto.vault_load(key, Path(self.backup_edit.text().strip()).read_bytes())
            doc = Document(self.path)
            rep = doc.decrypt(key, vault)
        except WrongKeyError as e:
            QMessageBox.warning(self, "Shifr", str(e))
            return
        except Exception as e:
            error_box(self, "Не удалось расшифровать", e)
            return
        default = str(output_path(self.path, ".dec"))
        out, _ = QFileDialog.getSaveFileName(self, "Сохранить расшифрованную презентацию", default, "PowerPoint (*.pptx)")
        if not out:
            return
        if Path(out).resolve() == self.path.resolve():
            QMessageBox.warning(self, "Shifr", "Нельзя перезаписывать исходный файл. Выберите другое имя.")
            return
        try:
            doc.save(out)
        except Exception as e:
            error_box(self, "Не удалось сохранить", e)
            return
        self._show_report(rep, Path(out))

    def _show_report(self, rep, out: Path):
        ok = rep.restored - sum(1 for d in rep.damaged if d["from_backup"])
        parts = [f"Сохранено: {out}", f"Восстановлено фрагментов: {rep.restored}."]
        if rep.damaged or rep.missing:
            parts.append(
                "Часть фрагментов требует ручной проверки — см. таблицу. "
                "Испорченные метки без резервной копии остаются на слайде как есть."
            )
        else:
            parts.append("Все найденные фрагменты расшифрованы без ошибок.")
        if not rep.missing and rep.damaged == [] and ok:
            parts.append("Если ИИ удалил последние по номеру фрагменты, программа без резервной копии этого не заметит — просмотрите результат.")
        self.summary.setText("\n".join(parts))

        rows = [(d["id"], d["where"], "восстановлен из резервной копии" if d["from_backup"]
                 else "метка испорчена, текст примерный", d["text"]) for d in rep.damaged]
        rows += [(m["id"], m["where"], "метка не найдена", m["text"] or "(нет резервной копии — текст неизвестен)")
                 for m in rep.missing]
        self.table.setRowCount(len(rows))
        self.table.setVisible(bool(rows))
        for r, (fid, where, state, text) in enumerate(rows):
            for c, val in enumerate((str(fid), where, state, text)):
                self.table.setItem(r, c, QTableWidgetItem(val))
            b = QPushButton("Скопировать")
            b.clicked.connect(lambda _=False, t=text: QGuiApplication.clipboard().setText(t))
            self.table.setCellWidget(r, 4, b)
        self.table.resizeColumnsToContents()
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)


# ============================================================== окно
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"Shifr {__version__} — шифрование текста в презентациях")
        self.resize(1150, 760)
        self.setAcceptDrops(True)
        self.tabs = QTabWidget()
        self.enc = EncryptTab()
        self.dec = DecryptTab()
        self.tabs.addTab(self.enc, "Зашифровать")
        self.tabs.addTab(self.dec, "Расшифровать")
        self.setCentralWidget(self.tabs)

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def dropEvent(self, e):
        for url in e.mimeData().urls():
            p = Path(url.toLocalFile())
            if p.suffix.lower() == ".pptx":
                (self.enc if self.tabs.currentIndex() == 0 else self.dec).load(p)
                return


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Shifr")
    w = MainWindow()
    w.show()
    if os.environ.get("SHIFR_SMOKE"):  # проверка сборки: окно создаётся и сразу закрывается
        app.processEvents()
        print("Shifr smoke OK", __version__)
        return
    if len(sys.argv) > 1 and Path(sys.argv[1]).suffix.lower() == ".pptx":
        w.enc.load(Path(sys.argv[1]))
    sys.exit(app.exec())
