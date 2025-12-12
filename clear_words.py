import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel
from tkinter import ttk
from tkinter import simpledialog
from collections import deque
import re
import webbrowser
import io
import urllib.request
import urllib.parse
import json
import zipfile
import os

# ===========================
#  ГЛОБАЛЬНЫЕ ПЕРЕМЕННЫЕ
# ===========================
all_data = None           # все данные (с колонкой _group)
filtered_data = None      # данные после стоп-слов
view_data = None          # данные с учетом поиска/фильтра по группе (то, что показываем в таблице)

stop_words = set()
history = deque()         # история стоп-слов (как было)

current_search_query = "" # строка поиска
current_group_filter = None  # None или имя группы
groups = []               # список веток (групп)
last_target_group = None  # последняя выбранная ветка для переноса
current_project_path = None  # путь к открытому проекту (.kwproj)

# ===========================
#  BUKVARIX API
# ===========================
BUKVARIX_API_URL_MKEYWORDS = "https://api.bukvarix.com/v1/mkeywords/"  # расширенный поиск (POST)

# ===========================
#  УТИЛИТЫ ЗАГРУЗКИ/СОХРАНЕНИЯ
# ===========================
def load_file():
    file_path = filedialog.askopenfilename(
        filetypes=[("CSV и Excel", "*.csv *.xlsx"), ("CSV files", "*.csv"), ("Excel files", "*.xlsx")]
    )
    if not file_path:
        return None
    try:
        if file_path.endswith(".csv"):
            try:
                return pd.read_csv(file_path, sep=';', encoding='utf-8')
            except pd.errors.ParserError:
                return pd.read_csv(file_path, sep=',', encoding='utf-8')
        else:
            return pd.read_excel(file_path)
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось загрузить файл:\n{e}")
        return None

def ensure_group_column(df: pd.DataFrame) -> pd.DataFrame:
    """Гарантируем служебную колонку '_group' (ветка) и убираем NaN."""
    if df is None:
        return df
    out = df.copy()

    # Если в файле уже есть пользовательская колонка "Группа" — используем её как источник
    if "_group" not in out.columns:
        if "Группа" in out.columns:
            out["_group"] = out["Группа"]
        else:
            out["_group"] = ""

    # Нормализуем: NaN -> пусто, строковые 'nan' тоже -> пусто
    out["_group"] = out["_group"].fillna("").astype(str)
    out["_group"] = out["_group"].replace({"nan": "", "None": ""})

    return out

def apply_stop_words_to_data(data, stop_words_set):
    if data is None:
        return None
    if not stop_words_set:
        return data.copy()

    # стоп-слова применяем к колонке "Фраза" (1-й видимый столбец)
    # У нас в all_data первый "логический" столбец — это "Фраза" (обычно 0-й индекс)
    phrase_col = data.columns[0]
    pattern = r'\b(' + '|'.join(map(re.escape, stop_words_set)) + r')\b'
    return data[~data[phrase_col].astype(str).str.contains(pattern, na=False, regex=True)].copy()

def rebuild_groups_list():
    """Перестраивает список веток из данных + локального списка groups (игнорируем NaN)."""
    global groups
    if all_data is None:
        groups = sorted(set([g for g in groups if str(g).strip() and str(g).lower() != "nan"]))
        return

    existing = set([g for g in groups if str(g).strip() and str(g).lower() != "nan"])

    vals = all_data.get("_group")
    if vals is not None:
        for g in vals.fillna("").astype(str).tolist():
            g = g.strip()
            if g and g.lower() != "nan":
                existing.add(g)

    groups = sorted(existing)

def refresh_groups_ui():
    # Перерисовка списка веток без потери текущего фильтра/выбора
    groups_list.delete(0, tk.END)
    groups_list.insert(tk.END, "Все")
    groups_list.insert(tk.END, "Без ветки")
    for g in groups:
        groups_list.insert(tk.END, g)

    # Восстанавливаем выделение под текущий фильтр
    # None -> "Все", "__NO_GROUP__" -> "Без ветки", иначе -> нужная ветка
    try:
        if current_group_filter is None:
            groups_list.selection_set(0)
            groups_list.activate(0)
        elif current_group_filter == "__NO_GROUP__":
            groups_list.selection_set(1)
            groups_list.activate(1)
        else:
            items = groups_list.get(0, tk.END)
            if current_group_filter in items:
                idx = items.index(current_group_filter)
                groups_list.selection_set(idx)
                groups_list.activate(idx)
    except Exception:
        pass

def update_view_data():
    """Применяем поиск + фильтр по ветке к filtered_data и получаем view_data."""
    global view_data
    if filtered_data is None:
        view_data = None
        return

    df = filtered_data.copy()

    # Фильтр по ветке
    if current_group_filter is not None:
        if current_group_filter == "__NO_GROUP__":
            df = df[df["_group"].fillna("").astype(str).str.strip() == ""]
        else:
            df = df[df["_group"].astype(str) == current_group_filter]

    # Поиск по словам (AND-логика)
    q = (current_search_query or "").strip().lower()
    if q:
        words = [w for w in q.split() if w]
        if words:
            phrase_col = df.columns[0]
            s = df[phrase_col].astype(str).str.lower()
            mask = pd.Series(True, index=df.index)
            for w in words:
                mask = mask & s.str.contains(re.escape(w), na=False, regex=True)
            df = df[mask]

    view_data = df

def update_filtered_data():
    global filtered_data
    if all_data is None:
        filtered_data = None
        return
    filtered_data = apply_stop_words_to_data(all_data, stop_words)
    update_view_data()

def refresh_table():
    # Очистка текущего содержимого Treeview
    for item in tree.get_children():
        tree.delete(item)

    if view_data is not None and not view_data.empty:
        # Показ: "Группа" + исходные колонки (кроме служебной _group)
        phrase_col = view_data.columns[0]
        for _, row in view_data.iterrows():
            values = []
            values.append("" if str(row.get("_group", "")).lower()=="nan" else row.get("_group", ""))
            # добавляем все колонки из исходника (кроме _group)
            for c in all_columns_no_group:
                values.append(row.get(c, ""))
            tree.insert("", "end", values=values)
    else:
        # Нет данных — показываем заглушку
        empty = [""] * (len(display_columns))
        if empty:
            empty[0] = ""
            empty[1] = "Данные отсутствуют"
        tree.insert("", "end", values=empty)

    count_label.config(text=f"Показано строк: {len(tree.get_children())}")

def add_stop_word(word):
    if word in stop_words:
        messagebox.showinfo("Информация", f"Слово '{word}' уже в списке стоп-слов.")
        return
    stop_words.add(word)
    history.append(stop_words.copy())
    update_filtered_data()
    refresh_table()

def remove_stop_word(word):
    if word in stop_words:
        stop_words.remove(word)
        history.append(stop_words.copy())
        update_filtered_data()
        refresh_table()

def open_word_selection(row_phrase):
    top = Toplevel(root)
    top.title("Выбор слова")
    top.geometry("320x240")

    words = row_phrase.split()
    unique_words = set(word.strip(",.!?;:\"'()[]{}") for word in words if word.strip(",.!?;:\"'()[]{}"))
    if not unique_words:
        messagebox.showinfo("Информация", "Не удалось выделить слова из фразы.")
        top.destroy()
        return

    for word in sorted(unique_words, key=lambda x: x.lower()):
        btn = tk.Button(
            top,
            text=f"Добавить '{word}' в стоп-слова",
            command=lambda w=word: [add_stop_word(w), top.destroy()]
        )
        btn.pack(pady=2, fill='x', padx=10)

def show_stop_words():
    top = Toplevel(root)
    top.title("Список стоп-слов")
    top.geometry("320x420")

    frame = tk.Frame(top)
    frame.pack(fill='both', expand=True, padx=10, pady=10)

    for word in sorted(stop_words, key=lambda x: x.lower()):
        row_frame = tk.Frame(frame)
        row_frame.pack(fill='x', pady=2)

        lbl = tk.Label(row_frame, text=word)
        lbl.pack(side='left', fill='x', expand=True)

        btn = tk.Button(
            row_frame,
            text="Удалить",
            command=lambda w=word: [remove_stop_word(w), top.destroy(), show_stop_words()]
        )
        btn.pack(side='right')

def undo_last_action():
    if history:
        last_state = history.pop()
        global stop_words
        stop_words = last_state
        update_filtered_data()
        refresh_table()
    else:
        messagebox.showinfo("Отмена", "Нет действий для отмены.")

def save_stop_words_to_file():
    if not stop_words:
        messagebox.showinfo("Информация", "Список стоп-слов пуст.")
        return
    file_path = filedialog.asksaveasfilename(
        defaultextension=".txt",
        filetypes=[("Text files", "*.txt")]
    )
    if file_path:
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write('\n'.join(sorted(stop_words, key=lambda x: x.lower())))
            messagebox.showinfo("Успех", "Список стоп-слов сохранен!")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{e}")

def load_stop_words_from_file():
    file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
    if file_path:
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                words = set(line.strip() for line in f if line.strip())
                if words:
                    history.append(stop_words.copy())
                    stop_words.update(words)
                    update_filtered_data()
                    refresh_table()
                    messagebox.showinfo("Успех", "Стоп-слова загружены и применены.")
                else:
                    messagebox.showinfo("Информация", "Файл стоп-слов пуст.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить файл:\n{e}")

def save_file():
    """
    Сохраняем filtered_data (после стоп-слов).

    XLSX:
      - каждая ветка (_group) = отдельный лист
      - отдельный лист "Стоп-слова"
    CSV:
      - один файл (листов в CSV быть не может), добавляется колонка "Группа"
    """
    if filtered_data is None or filtered_data.empty:
        messagebox.showinfo("Информация", "Нет данных для сохранения.")
        return

    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
    )
    if not file_path:
        return

    try:
        if file_path.lower().endswith(".csv"):
            out = filtered_data.copy()
            if "_group" in out.columns:
                out.insert(0, "Группа", out["_group"])
                out = out.drop(columns=["_group"])
            out.to_csv(file_path, index=False)
            messagebox.showinfo("Успех", "CSV успешно сохранён!")
            return

        # XLSX: пишем через openpyxl — так Excel точно увидит листы
        from openpyxl import Workbook
        from openpyxl.utils.dataframe import dataframe_to_rows

        df = filtered_data.copy()
        if "_group" not in df.columns:
            df["_group"] = ""

        def _sheet_name(raw: str) -> str:
            name = (raw or "").strip()
            name = re.sub(r"[\[\]\:\*\?\/\\]", "_", name)
            name = name.strip()
            if not name:
                name = "Без ветки"
            if len(name) > 31:
                name = name[:31]
            return name

        def _unique_sheet_name(name: str, used: set) -> str:
            base = name
            if base not in used:
                used.add(base)
                return base
            i = 2
            while True:
                suffix = f"_{i}"
                candidate = base
                if len(candidate) + len(suffix) > 31:
                    candidate = candidate[:31 - len(suffix)]
                candidate = candidate + suffix
                if candidate not in used:
                    used.add(candidate)
                    return candidate
                i += 1

        wb = Workbook()
        # удаляем дефолтный лист "Sheet"
        if wb.worksheets:
            wb.remove(wb.worksheets[0])

        used = set()

        # список групп (включая пустую)
        groups_unique = df["_group"].astype(str).fillna("").tolist()
        groups_unique = sorted(set([g.strip() for g in groups_unique]))

        if "" in groups_unique:
            groups_unique.remove("")
            ordered = [""] + groups_unique
        else:
            ordered = groups_unique

        # листы по веткам
        for g in ordered:
            part = df[df["_group"].astype(str).fillna("").str.strip() == g].copy()
            part = part.drop(columns=["_group"], errors="ignore")

            sheet = _unique_sheet_name(_sheet_name(g), used)
            ws = wb.create_sheet(title=sheet)

            for r_idx, row in enumerate(dataframe_to_rows(part, index=False, header=True), start=1):
                ws.append(row)

        # лист стоп-слов (даже если пусто — лист должен быть)
        sw = sorted(stop_words, key=lambda x: x.lower())
        sw_df = pd.DataFrame({"Стоп-слова": sw})
        sw_sheet = _unique_sheet_name(_sheet_name("Стоп-слова"), used)
        ws_sw = wb.create_sheet(title=sw_sheet)
        for row in dataframe_to_rows(sw_df, index=False, header=True):
            ws_sw.append(row)

        wb.save(file_path)

        messagebox.showinfo("Успех", f"Excel сохранён.\nЛистов: {len(wb.worksheets)} (ветки + стоп-слова).")

    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{e}")


def load_data():
    global all_data, filtered_data, view_data, current_search_query, current_group_filter
    df = load_file()
    if df is None:
        return
    all_data = ensure_group_column(df)
    current_search_query = ""
    current_group_filter = None
    search_var.set("")
    rebuild_groups_list()
    refresh_groups_ui()
    update_filtered_data()
    refresh_table()

# ===========================
#  СОРТИРОВКИ
# ===========================
def sort_by_column(index_in_original, ascending=True):
    """
    Сортируем filtered_data по колонке исходных данных (без служебной _group).
    index_in_original — индекс относительно all_columns_no_group.
    """
    global filtered_data
    if filtered_data is None or filtered_data.empty:
        return

    col = all_columns_no_group[index_in_original]
    try:
        filtered_data = filtered_data.sort_values(by=col, ascending=ascending, kind="mergesort")
    except Exception:
        filtered_data = filtered_data.sort_values(by=col, ascending=ascending, kind="mergesort")

    update_view_data()
    refresh_table()

def sort_alphabetically():
    sort_by_column(0, ascending=True)

def sort_by_statistics1():
    if len(all_columns_no_group) >= 2:
        sort_by_column(1, ascending=False)

def sort_by_statistics2():
    if len(all_columns_no_group) >= 3:
        sort_by_column(2, ascending=False)

def sort_by_statistics3():
    if len(all_columns_no_group) >= 4:
        sort_by_column(3, ascending=False)

# ===========================
#  ДЕДУПЛИКАЦИЯ
# ===========================
def _normalize_phrase_exact(value: str) -> str:
    s = str(value).lower().replace("ё", "е").strip()
    s = re.sub(r"\s+", " ", s)
    return s

def _normalize_phrase_tokensort(value: str) -> str:
    s = str(value).lower().replace("ё", "е")
    tokens = re.findall(r"[0-9a-zа-я]+", s, flags=re.IGNORECASE)
    tokens = [t for t in tokens if t]
    tokens.sort()
    return " ".join(tokens)

def _dedupe_keep_most_frequent(df: pd.DataFrame, key_func) -> pd.DataFrame:
    if df is None or df.empty:
        return df

    tmp = df.copy()

    # Ключ дедупликации строим по фразе (первая исходная колонка)
    phrase_col = all_columns_no_group[0]
    tmp["_dedupe_key"] = tmp[phrase_col].astype(str).map(key_func)

    # Колонка "частотности" — ищем типовую, иначе берём второй столбец
    freq_col = None
    if "Частотность" in tmp.columns:
        freq_col = "Частотность"
    elif len(all_columns_no_group) >= 2:
        freq_col = all_columns_no_group[1]

    if freq_col is not None:
        tmp["_freq1"] = pd.to_numeric(tmp[freq_col], errors="coerce").fillna(-1)

    if len(all_columns_no_group) >= 3:
        tmp["_freq2"] = pd.to_numeric(tmp[all_columns_no_group[2]], errors="coerce").fillna(-1)

    tmp["_len"] = tmp[phrase_col].astype(str).str.len()

    sort_cols = ["_dedupe_key"]
    ascending = [True]

    if "_freq1" in tmp.columns:
        sort_cols.append("_freq1")
        ascending.append(False)

    if "_freq2" in tmp.columns:
        sort_cols.append("_freq2")
        ascending.append(False)

    sort_cols.append("_len")
    ascending.append(True)

    tmp = tmp.sort_values(by=sort_cols, ascending=ascending, kind="mergesort")
    tmp = tmp.drop_duplicates(subset=["_dedupe_key"], keep="first")

    for c in ["_dedupe_key", "_freq1", "_freq2", "_len"]:
        if c in tmp.columns:
            tmp = tmp.drop(columns=[c])

    return tmp

def remove_duplicates_exact():
    global filtered_data
    if filtered_data is None or filtered_data.empty:
        messagebox.showinfo("Информация", "Нет данных для дедупликации.")
        return

    before = len(filtered_data)
    filtered_data = _dedupe_keep_most_frequent(filtered_data, _normalize_phrase_exact)
    after = len(filtered_data)
    rebuild_groups_list()
    refresh_groups_ui()
    update_view_data()
    refresh_table()

    messagebox.showinfo(
        "Готово",
        f"Удалены явные дубли: {before - after}\nОсталось строк: {after}\n\n"
        "Правило: сравнение по фразе без учета регистра/лишних пробелов/ё."
    )

def remove_duplicates_soft():
    global filtered_data
    if filtered_data is None or filtered_data.empty:
        messagebox.showinfo("Информация", "Нет данных для дедупликации.")
        return

    before = len(filtered_data)
    filtered_data = _dedupe_keep_most_frequent(filtered_data, _normalize_phrase_tokensort)
    after = len(filtered_data)
    rebuild_groups_list()
    refresh_groups_ui()
    update_view_data()
    refresh_table()

    messagebox.showinfo(
        "Готово",
        f"Удалены неявные дубли: {before - after}\nОсталось строк: {after}\n\n"
        "Правило: игнорируется пунктуация и порядок слов (слова сортируются)."
    )

# ===========================
#  ГРУППИРОВКА (ВЕТКИ)
# ===========================
def add_group():
    name = simpledialog.askstring("Новая ветка", "Введите название ветки (например: Цена, Город):")
    if not name:
        return
    name = name.strip()
    if not name:
        return
    if name in groups:
        messagebox.showinfo("Информация", f"Ветка '{name}' уже существует.")
        return
    groups.append(name)
    rebuild_groups_list()
    refresh_groups_ui()

def delete_group():
    sel = groups_list.curselection()
    if not sel:
        messagebox.showinfo("Информация", "Выберите ветку справа.")
        return
    label = groups_list.get(sel[0])
    if label in ("Все", "Без ветки"):
        messagebox.showinfo("Информация", "Эту системную ветку удалить нельзя.")
        return

    if not messagebox.askyesno("Удалить ветку", f"Удалить ветку '{label}'?\n\nКлючи из этой ветки станут 'Без ветки'."):
        return

    global all_data
    if all_data is not None:
        all_data.loc[all_data["_group"].astype(str) == label, "_group"] = ""

    if label in groups:
        groups.remove(label)

    # Если сейчас фильтруем по этой ветке — сбросим фильтр
    global current_group_filter
    if current_group_filter == label:
        current_group_filter = None

    rebuild_groups_list()
    refresh_groups_ui()
    update_filtered_data()
    refresh_table()

def _pick_target_group(title="Выбор ветки", hint="Выберите ветку для переноса:"):
    """Если справа выбран 'Все' — спрашиваем целевую ветку, чтобы не ломать UX."""
    global last_target_group
    if not groups:
        messagebox.showinfo("Информация", "Сначала создайте ветку кнопкой '+ Ветка'.")
        return None

    top = Toplevel(root)
    top.title(title)
    top.geometry("320x360")
    top.transient(root)
    top.grab_set()

    tk.Label(top, text=hint, anchor="w", justify="left").pack(fill="x", padx=10, pady=(10, 5))

    lb = tk.Listbox(top, height=12, exportselection=False)
    lb.pack(fill="both", expand=True, padx=10)

    lb.insert(tk.END, "Без ветки")
    for g in groups:
        lb.insert(tk.END, g)

    # выделение по умолчанию
    items = lb.get(0, tk.END)
    if last_target_group:
        if last_target_group == "":
            lb.selection_set(0)
            lb.activate(0)
        elif last_target_group in items:
            lb.selection_set(items.index(last_target_group))
            lb.activate(items.index(last_target_group))
        else:
            lb.selection_set(1)
            lb.activate(1)
    else:
        lb.selection_set(1)
        lb.activate(1)

    result = {"value": None}

    def _ok():
        sel = lb.curselection()
        if not sel:
            messagebox.showinfo("Информация", "Выберите ветку.")
            return
        label = lb.get(sel[0])
        if label == "Без ветки":
            result["value"] = ""
            last_target_group = ""
        else:
            result["value"] = label
            last_target_group = label
        top.destroy()

    def _cancel():
        result["value"] = None
        top.destroy()

    btns = tk.Frame(top)
    btns.pack(fill="x", padx=10, pady=10)
    tk.Button(btns, text="OK", command=_ok).pack(side="left")
    tk.Button(btns, text="Отмена", command=_cancel).pack(side="right")

    top.wait_window()
    return result["value"]

def set_group_filter_from_list(event=None):
    global current_group_filter
    sel = groups_list.curselection()
    if not sel:
        return
    label = groups_list.get(sel[0])

    if label == "Все":
        current_group_filter = None
    elif label == "Без ветки":
        current_group_filter = "__NO_GROUP__"
    else:
        current_group_filter = label

    update_view_data()
    refresh_table()

def move_selected_to_group():
    if all_data is None or filtered_data is None:
        messagebox.showinfo("Информация", "Сначала загрузите данные.")
        return

    sel_items = tree.selection()
    if not sel_items:
        messagebox.showinfo("Информация", "Выделите строки в таблице.")
        return
    # целевая ветка
    sel_group = groups_list.curselection()
    target = None

    if sel_group:
        target_label = groups_list.get(sel_group[0])
        if target_label == "Без ветки":
            target = ""
        elif target_label == "Все":
            target = _pick_target_group()
        else:
            target = target_label
    else:
        # если ничего не выбрано — спросим
        target = _pick_target_group()

    if target is None:
        return

    # Фраза в Treeview теперь во 2-й колонке (после "Группа")
    phrases = []
    for item in sel_items:
        vals = tree.item(item, "values")
        if not vals:
            continue
        # vals[1] — это "Фраза" (первая исходная колонка)
        if len(vals) >= 2 and vals[1] != "Данные отсутствуют":
            phrases.append(vals[1])

    if not phrases:
        return

    phrase_col = all_columns_no_group[0]
    mask = all_data[phrase_col].astype(str).isin([str(p) for p in phrases])
    all_data.loc[mask, "_group"] = target

    rebuild_groups_list()
    refresh_groups_ui()
    update_filtered_data()
    refresh_table()

def move_all_view_to_group():
    if all_data is None or view_data is None or view_data.empty:
        messagebox.showinfo("Информация", "Нет данных для переноса (таблица пустая).")
        return

    sel_group = groups_list.curselection()
    if not sel_group:
        messagebox.showinfo("Информация", "Выберите ветку справа (или создайте новую).")
        return
    target_label = groups_list.get(sel_group[0])
    if target_label in ("Все",):
        messagebox.showinfo("Информация", "Выберите конкретную ветку справа (не 'Все').")
        return
    if target_label == "Без ветки":
        target = ""
    else:
        target = target_label

    phrase_col = all_columns_no_group[0]
    phrases = view_data[phrase_col].astype(str).tolist()
    mask = all_data[phrase_col].astype(str).isin(phrases)
    all_data.loc[mask, "_group"] = target

    rebuild_groups_list()
    refresh_groups_ui()
    update_filtered_data()
    refresh_table()


# ===========================
#  УДАЛЕНИЕ СТРОК (ПКМ)
# ===========================
def delete_selected_rows():
    """
    Удаляет выбранные строки из all_data (т.е. полностью из проекта),
    затем пересчитывает filtered_data/view_data и обновляет таблицу.
    """
    global all_data
    if all_data is None:
        return

    sel_items = tree.selection()
    if not sel_items:
        messagebox.showinfo("Информация", "Выделите строки в таблице для удаления.")
        return

    phrases = []
    for item in sel_items:
        vals = tree.item(item, "values")
        if vals and len(vals) >= 2 and vals[1] != "Данные отсутствуют":
            phrases.append(str(vals[1]))

    if not phrases:
        return

    if not messagebox.askyesno("Удаление", f"Удалить выбранные строки: {len(phrases)}?\nЭто действие удалит фразы из набора данных."):
        return

    phrase_col = all_columns_no_group[0]
    all_data = all_data[~all_data[phrase_col].astype(str).isin(phrases)].copy()

    rebuild_groups_list()
    refresh_groups_ui()
    update_filtered_data()
    refresh_table()

def delete_row_under_cursor(event):
    """
    ПКМ по строке: если строка не выделена — выделяем её, затем показываем меню.
    """
    item = tree.identify_row(event.y)
    if item:
        if item not in tree.selection():
            tree.selection_set(item)
            tree.focus(item)
        try:
            tree_menu.tk_popup(event.x_root, event.y_root)
        finally:
            tree_menu.grab_release()

# ===========================
#  DRAG & DROP: ПЕРЕТЯГИВАНИЕ СТРОК В ВЕТКУ
# ===========================
_drag_state = {"active": False, "phrases": []}

_drag_state["press_xy"] = (0, 0)

def _widget_is_or_inside(widget, target) -> bool:
    w = widget
    while w is not None:
        if w == target:
            return True
        w = getattr(w, "master", None)
    return False

def on_tree_mouse_down(event):
    """
    Стартуем drag так, чтобы НЕ сбрасывалось мультивыделение.
    Особый случай: кликаем по уже выделенной строке при множественном выделении —
    гасим дефолтное поведение Treeview (которое сбрасывает выделение).
    """
    item = tree.identify_row(event.y)
    if not item:
        return

    current_sel = set(tree.selection())
    _drag_state["press_xy"] = (event.x_root, event.y_root)

    # Если клик по уже выделенной строке и выделено несколько — не даём Treeview сбросить выделение
    if item in current_sel and len(current_sel) > 1:
        phrases = _get_selected_phrases_from_tree()
        if not phrases:
            return "break"
        _drag_state["active"] = True
        _drag_state["phrases"] = phrases
        root.config(cursor="hand2")
        return "break"

    # Иначе — даём Treeview обработать клик/выделение, а drag-данные возьмём после этого
    def _late_start():
        phrases = _get_selected_phrases_from_tree()
        if not phrases:
            return
        _drag_state["active"] = True
        _drag_state["phrases"] = phrases
        root.config(cursor="hand2")
    root.after(1, _late_start)
    # не возвращаем break

def on_tree_mouse_up(event):
    """
    Если отпустили кнопку над списком веток — переносим выбранные фразы в ветку под курсором.
    """
    global all_data
    if not _drag_state.get("active"):
        return

    _drag_state["active"] = False
    root.config(cursor="")

    if all_data is None:
        return

    w = root.winfo_containing(event.x_root, event.y_root)
    if not _widget_is_or_inside(w, groups_list):
        return

    # Рассчитываем индекс элемента listbox, над которым отпустили
    y_local = event.y_root - groups_list.winfo_rooty()
    idx = groups_list.nearest(y_local)
    if idx is None:
        return

    label = groups_list.get(idx)

    if label == "Все":
        target = _pick_target_group(title="Перенос в ветку", hint="Выберите ветку, в которую перенести выбранные фразы:")
        if target is None:
            return
    elif label == "Без ветки":
        target = ""
    else:
        target = label

    phrases = _drag_state.get("phrases") or []
    if not phrases:
        return

    phrase_col = all_columns_no_group[0]
    mask = all_data[phrase_col].astype(str).isin(phrases)
    all_data.loc[mask, "_group"] = target

    rebuild_groups_list()
    refresh_groups_ui()
    update_filtered_data()
    refresh_table()

def _get_selected_phrases_from_tree():
    sel_items = tree.selection()
    phrases = []
    for item in sel_items:
        vals = tree.item(item, "values")
        if vals and len(vals) >= 2 and vals[1] != "Данные отсутствуют":
            phrases.append(str(vals[1]))
    return phrases


    """
    Если отпустили кнопку над списком веток — переносим выбранные фразы в ветку под курсором.
    """
    global all_data
    if not _drag_state["active"]:
        return

    _drag_state["active"] = False
    root.config(cursor="")

    if all_data is None:
        return

    # Определяем, над каким виджетом отпустили
    w = root.winfo_containing(event.x_root, event.y_root)
    if w != groups_list:
        return

    idx = groups_list.nearest(event.y_root - groups_list.winfo_rooty())
    if idx is None:
        return

    label = groups_list.get(idx)

    if label == "Все":
        target = _pick_target_group(title="Перенос в ветку", hint="Выберите ветку, в которую перенести выбранные фразы:")
        if target is None:
            return
    elif label == "Без ветки":
        target = ""
    else:
        target = label

    phrases = _drag_state.get("phrases") or []
    if not phrases:
        return

    phrase_col = all_columns_no_group[0]
    mask = all_data[phrase_col].astype(str).isin(phrases)
    all_data.loc[mask, "_group"] = target

    rebuild_groups_list()
    refresh_groups_ui()
    update_filtered_data()
    refresh_table()
# ===========================
#  ПОИСК
# ===========================
def on_search_change(*args):
    global current_search_query
    current_search_query = search_var.get()
    update_view_data()
    refresh_table()

def clear_search():
    search_var.set("")

# ===========================
#  BUKVARIX
# ===========================
def contact_author():
    top = Toplevel(root)
    top.title("Связаться с автором")
    top.geometry("300x200")

    lbl = tk.Label(top, text="Обработка слов v2.0\nАвтор: Эльдар Ибрагимов", justify=tk.CENTER)
    lbl.pack(pady=10)

    btn_vk = tk.Button(top, text="ВК", command=lambda: webbrowser.open("https://vk.com/mr.crutch"))
    btn_vk.pack(side=tk.LEFT, padx=30)

    btn_tg = tk.Button(top, text="TG", command=lambda: webbrowser.open("https://t.me/God_SMM"))
    btn_tg.pack(side=tk.RIGHT, padx=30)

def _bukvarix_request_mkeywords(keywords, api_key="free", num=250, timeout_sec=45):
    if not keywords:
        raise ValueError("Список ключевых слов пуст.")
    payload = {
        "api_key": (api_key or "free").strip(),
        "q": "\r\n".join(keywords),
        "format": "csv",
        "header": "0",
        "num": str(int(num)),
        "bom": "1",
    }
    data = urllib.parse.urlencode(payload).encode("utf-8")
    req = urllib.request.Request(
        BUKVARIX_API_URL_MKEYWORDS,
        data=data,
        method="POST",
        headers={"Content-Type": "application/x-www-form-urlencoded"},
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout_sec) as resp:
            raw = resp.read()
    except urllib.error.HTTPError as e:
        body = ""
        try:
            body = e.read().decode("utf-8", errors="replace")
        except Exception:
            pass
        raise RuntimeError(f"Bukvarix вернул HTTP {e.code}: {body or e.reason}") from e
    except Exception as e:
        raise RuntimeError(f"Не удалось выполнить запрос к Bukvarix: {e}") from e

    csv_text = raw.decode("utf-8-sig", errors="replace")

    df = pd.read_csv(
        io.StringIO(csv_text),
        sep=";",
        header=None,
        names=["phrase", "words", "symbols", "broad", "exact"],
    )
    return df

def open_bukvarix_parser():
    top = Toplevel(root)
    top.title("Парсинг Bukvarix")
    top.geometry("620x420")

    lbl = tk.Label(
        top,
        text="Введите до 100 ключевых слов/фраз (каждое с новой строки):",
        anchor="w",
        justify=tk.LEFT,
    )
    lbl.pack(fill="x", padx=10, pady=(10, 5))

    txt = tk.Text(top, height=12)
    txt.pack(fill="both", expand=True, padx=10)

    opts = tk.Frame(top)
    opts.pack(fill="x", padx=10, pady=10)

    tk.Label(opts, text="API key:").grid(row=0, column=0, sticky="w")
    api_key_var = tk.StringVar(value="free")
    api_entry = tk.Entry(opts, textvariable=api_key_var, width=25)
    api_entry.grid(row=0, column=1, sticky="w", padx=(5, 20))

    tk.Label(opts, text="Строк в отчете (num):").grid(row=0, column=2, sticky="w")
    num_var = tk.IntVar(value=250)
    num_spin = tk.Spinbox(opts, from_=10, to=1000000, textvariable=num_var, width=10)
    num_spin.grid(row=0, column=3, sticky="w", padx=(5, 0))

    btns = tk.Frame(top)
    btns.pack(fill="x", padx=10, pady=(0, 10))

    def _run_parse():
        raw_input = txt.get("1.0", "end").splitlines()
        seeds = [line.strip() for line in raw_input if line.strip()]
        if not seeds:
            messagebox.showinfo("Информация", "Введите хотя бы одно ключевое слово.")
            return
        if len(seeds) > 100:
            messagebox.showerror("Ошибка", "Можно ввести не более 100 ключевых слов за раз.")
            return

        root.config(cursor="watch")
        top.config(cursor="watch")
        top.update_idletasks()

        try:
            df_raw = _bukvarix_request_mkeywords(
                seeds,
                api_key=api_key_var.get().strip() or "free",
                num=num_var.get(),
            )
            if df_raw is None or df_raw.empty:
                messagebox.showinfo("Результат", "Bukvarix не вернул данные по заданным ключам.")
                return

            df_gui = pd.DataFrame({
                "Фраза": df_raw["phrase"].astype(str),
                "Частотность": pd.to_numeric(df_raw["broad"], errors="coerce"),
                "!Частостность": pd.to_numeric(df_raw["exact"], errors="coerce"),
                "[!Частостность]": pd.to_numeric(df_raw["exact"], errors="coerce"),
            })

            global all_data, current_search_query, current_group_filter
            all_data = ensure_group_column(df_gui)
            current_search_query = ""
            current_group_filter = None
            search_var.set("")

            rebuild_groups_list()
            refresh_groups_ui()
            update_filtered_data()
            refresh_table()

            messagebox.showinfo("Успех", f"Загружено строк из Bukvarix: {len(df_gui)}")
            top.destroy()

        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
        finally:
            root.config(cursor="")
            top.config(cursor="")

    btn_parse = tk.Button(btns, text="Парсить", command=_run_parse)
    btn_parse.pack(side="left")

    btn_cancel = tk.Button(btns, text="Отмена", command=top.destroy)
    btn_cancel.pack(side="right")


# ===========================
#  ПРОЕКТ: СОХРАНИТЬ / ОТКРЫТЬ (как в KeyCollector)
# ===========================
PROJECT_EXT = ".el"  # архив ZIP с данными и настройками

def _project_default_name():
    return "project" + PROJECT_EXT

def save_project_as():
    global current_project_path
    path = filedialog.asksaveasfilename(
        defaultextension=PROJECT_EXT,
        filetypes=[(f"Проект ({PROJECT_EXT})", f"*{PROJECT_EXT}")]
    )
    if not path:
        return
    if not path.lower().endswith(PROJECT_EXT):
        path += PROJECT_EXT
    current_project_path = path
    _save_project_to_path(current_project_path)

def save_project():
    global current_project_path
    if current_project_path:
        _save_project_to_path(current_project_path)
    else:
        save_project_as()

def _save_project_to_path(path: str):
    if all_data is None:
        messagebox.showinfo("Информация", "Нет данных для сохранения проекта. Сначала загрузите файл/парсинг.")
        return
    try:
        meta = {
            "version": 1,
            "groups": groups,
            "stop_words": sorted(stop_words, key=lambda x: x.lower()),
            "current_search_query": current_search_query,
            "current_group_filter": current_group_filter,
            "last_target_group": last_target_group,
            "columns": [c for c in all_data.columns.tolist()],
        }

        # сохраняем данные как CSV внутри ZIP (utf-8)
        df_bytes = all_data.to_csv(index=False).encode("utf-8")

        with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as z:
            z.writestr("data.csv", df_bytes)
            z.writestr("meta.json", json.dumps(meta, ensure_ascii=False, indent=2).encode("utf-8"))

        messagebox.showinfo("Успех", f"Проект сохранён:\n{path}")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось сохранить проект:\n{e}")

def open_project():
    global current_project_path, all_data, stop_words, groups
    global current_search_query, current_group_filter, last_target_group

    path = filedialog.askopenfilename(
        filetypes=[(f"Проект ({PROJECT_EXT})", f"*{PROJECT_EXT}")]
    )
    if not path:
        return

    try:
        with zipfile.ZipFile(path, "r") as z:
            if "data.csv" not in z.namelist() or "meta.json" not in z.namelist():
                raise ValueError("Неверный формат проекта: нет data.csv или meta.json")

            data_csv = z.read("data.csv").decode("utf-8", errors="replace")
            meta = json.loads(z.read("meta.json").decode("utf-8", errors="replace"))

        df = pd.read_csv(io.StringIO(data_csv))
        df = ensure_group_column(df)  # на всякий случай

        all_data = df

        # восстановление метаданных
        groups = meta.get("groups", [])
        stop_words = set(meta.get("stop_words", []))
        current_search_query = meta.get("current_search_query", "") or ""
        current_group_filter = meta.get("current_group_filter", None)
        last_target_group = meta.get("last_target_group", None)

        # UI: перестроение колонок/веток/поиска
        rebuild_columns_from_data()
        rebuild_groups_list()
        refresh_groups_ui()

        search_var.set(current_search_query)

        update_filtered_data()
        refresh_table()

        current_project_path = path
        messagebox.showinfo("Успех", f"Проект открыт:\n{path}")

    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось открыть проект:\n{e}")

# ===========================
#  UI
# ===========================
root = tk.Tk()
root.title("Обработка ключевых слов")
root.geometry("1600x780")

# Верхняя панель кнопок
btn_frame = tk.Frame(root)
btn_frame.pack(fill="x", pady=10, padx=10)

btn_load = tk.Button(btn_frame, text="Загрузить файл", command=load_data)
btn_load.pack(side="left", padx=5)

btn_open_project = tk.Button(btn_frame, text="Открыть проект", command=open_project)
btn_open_project.pack(side="left", padx=5)

btn_save_project = tk.Button(btn_frame, text="Сохранить проект", command=save_project)
btn_save_project.pack(side="left", padx=5)

btn_save_project_as = tk.Button(btn_frame, text="Сохранить проект как", command=save_project_as)
btn_save_project_as.pack(side="left", padx=5)

btn_save = tk.Button(btn_frame, text="Выгрузить файл", command=save_file)
btn_save.pack(side="left", padx=5)

btn_save_stop_words = tk.Button(btn_frame, text="Сохранить стоп-слова", command=save_stop_words_to_file)
btn_save_stop_words.pack(side="left", padx=5)

btn_load_stop_words = tk.Button(btn_frame, text="Загрузить стоп-слова", command=load_stop_words_from_file)
btn_load_stop_words.pack(side="left", padx=5)

btn_show_stop_words = tk.Button(btn_frame, text="Просмотр стоп-слов", command=show_stop_words)
btn_show_stop_words.pack(side="left", padx=5)

btn_undo = tk.Button(btn_frame, text="Назад", command=undo_last_action)
btn_undo.pack(side="left", padx=5)

# Меню сортировки/дублей
btn_sort = tk.Menubutton(btn_frame, text="Сортировка/Дубли", relief=tk.RAISED)
btn_sort.pack(side="left", padx=5)

sort_menu = tk.Menu(btn_sort, tearoff=0)
btn_sort.config(menu=sort_menu)
sort_menu.add_command(label="По алфавиту", command=sort_alphabetically)
sort_menu.add_command(label="Частотность", command=sort_by_statistics1)
sort_menu.add_command(label="\"!Частостность\"", command=sort_by_statistics2)
sort_menu.add_command(label="\"[!Частостность]\"", command=sort_by_statistics3)
sort_menu.add_separator()
sort_menu.add_command(label="Удалить дубли (явные, оставить более частотный)", command=remove_duplicates_exact)
sort_menu.add_command(label="Удалить дубли (неявные, по словам/пунктуации)", command=remove_duplicates_soft)

btn_bukvarix = tk.Button(btn_frame, text="Парсинг Bukvarix", command=open_bukvarix_parser)
btn_bukvarix.pack(side="left", padx=5)

btn_contact = tk.Button(btn_frame, text="Связаться с автором", command=contact_author)
btn_contact.pack(side="right", padx=5)

# Панель поиска
search_frame = tk.Frame(root)
search_frame.pack(fill="x", padx=10, pady=(0, 10))

tk.Label(search_frame, text="Поиск (AND по словам):").pack(side="left")

search_var = tk.StringVar()
search_var.trace_add("write", on_search_change)

search_entry = tk.Entry(search_frame, textvariable=search_var, width=60)
search_entry.pack(side="left", padx=8)

btn_clear_search = tk.Button(search_frame, text="Очистить", command=clear_search)
btn_clear_search.pack(side="left")

btn_move_selected = tk.Button(search_frame, text="Перенести выделенные → ветка", command=move_selected_to_group)
btn_move_selected.pack(side="right", padx=5)

btn_move_all = tk.Button(search_frame, text="Перенести ВСЕ из таблицы → ветка", command=move_all_view_to_group)
btn_move_all.pack(side="right", padx=5)

# Основной контейнер: слева таблица, справа ветки
main = tk.PanedWindow(root, orient=tk.HORIZONTAL, sashrelief=tk.RAISED)
main.pack(fill="both", expand=True, padx=10, pady=10)

# Левая часть: таблица
left = tk.Frame(main)
main.add(left, stretch="always")

tree_frame = tk.Frame(left)
tree_frame.pack(fill='both', expand=True)

# Правая часть: ветки
right = tk.Frame(main, width=240)
main.add(right)

tk.Label(right, text="Ветки (группировка)").pack(anchor="w", pady=(0, 5))

groups_list = tk.Listbox(right, height=18, exportselection=False)
groups_list.pack(fill="both", expand=False)
groups_list.bind("<<ListboxSelect>>", set_group_filter_from_list)

grp_btns = tk.Frame(right)
grp_btns.pack(fill="x", pady=8)

btn_add_group = tk.Button(grp_btns, text="+ Ветка", command=add_group)
btn_add_group.pack(side="left", padx=3)

btn_del_group = tk.Button(grp_btns, text="Удалить", command=delete_group)
btn_del_group.pack(side="left", padx=3)

tk.Label(
    right,
    text="Как работать:\n1) Создай ветку справа\n2) Найди ключи (поиск)\n3) Выдели строки или используй 'перенести все'\n4) Выгрузи файл — колонка 'Группа' будет в выгрузке",
    justify="left",
    wraplength=220
).pack(anchor="w", pady=10)

# ===========================
#  ТАБЛИЦА (Treeview)
# ===========================
# По умолчанию ориентируемся на классические 4 колонки.
# При загрузке будем строить "исходные" колонки динамически.
default_cols = ["Фраза", "Частотность", "!Частостность", "[!Частостность]"]
all_columns_no_group = default_cols.copy()

display_columns = ["Группа"] + all_columns_no_group  # то, что показываем в Treeview

tree = ttk.Treeview(tree_frame, columns=display_columns, show='headings', selectmode="extended")

for col in display_columns:
    tree.heading(col, text=col)
    w = 170
    if col == "Группа":
        w = 120
    elif col == "Фраза":
        w = 520
    tree.column(col, width=w, anchor='w')

vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
tree.configure(yscrollcommand=vsb.set)
vsb.pack(side='right', fill='y')

hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
tree.configure(xscrollcommand=hsb.set)
hsb.pack(side='bottom', fill='x')

tree.pack(fill='both', expand=True)

count_label = tk.Label(root, text="Показано строк: 0")
count_label.pack(side='bottom', padx=10, pady=5)

def rebuild_columns_from_data():
    """Если пользователь загрузил файл с другими колонками — перестроим отображение."""
    global all_columns_no_group, display_columns, tree, view_data

    if all_data is None:
        return

    # Берем все колонки, кроме служебной _group
    cols = [c for c in all_data.columns.tolist() if c != "_group"]
    if not cols:
        cols = default_cols.copy()

    all_columns_no_group = cols
    display_columns = ["Группа"] + all_columns_no_group

    # Пересобираем Treeview колонки
    tree["columns"] = display_columns
    for col in display_columns:
        tree.heading(col, text=col)
        w = 170
        if col == "Группа":
            w = 120
        elif col.lower() == "фраза" or col == cols[0]:
            w = 520
        tree.column(col, width=w, anchor='w')

def load_data():
    global all_data, filtered_data, view_data, current_search_query, current_group_filter
    df = load_file()
    if df is None:
        return
    all_data = ensure_group_column(df)
    current_search_query = ""
    current_group_filter = None
    search_var.set("")

    rebuild_columns_from_data()

    rebuild_groups_list()
    refresh_groups_ui()

    update_filtered_data()
    refresh_table()

# Переназначили load_data (кнопка уже связана), обновим command
btn_load.config(command=load_data)

def on_double_click(event):
    item = tree.identify_row(event.y)
    if item:
        values = tree.item(item, "values")
        if values and len(values) >= 2 and values[1] != "Данные отсутствуют":
            # values[1] — фраза (первая исходная колонка)
            open_word_selection(values[1])

tree.bind("<Double-1>", on_double_click)

# ПКМ-меню по таблице
tree_menu = tk.Menu(root, tearoff=0)
tree_menu.add_command(label="Удалить выбранное", command=delete_selected_rows)

tree.bind("<Button-3>", delete_row_under_cursor)   # ПКМ
# Drag & Drop в ветки (сохранение мультивыделения)
tree.bind("<ButtonPress-1>", on_tree_mouse_down, add="+")
tree.bind("<ButtonRelease-1>", on_tree_mouse_up, add="+")
# Drag & Drop в ветки

# Инициализация списка веток по умолчанию
refresh_groups_ui()

root.mainloop()
