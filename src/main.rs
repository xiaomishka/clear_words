#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use std::cmp::Ordering;
use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};
use std::error::Error;
use std::fs::File;
use std::io::{Read, Write};
use std::path::{Path, PathBuf};
use std::sync::mpsc::{self, Receiver, TryRecvError};
use std::time::Duration;

use calamine::{open_workbook_auto, Data, Reader};
use eframe::egui;
use egui_extras::{Column, TableBuilder};
use regex::Regex;
use rust_stemmers::{Algorithm, Stemmer};
use rfd::{FileDialog, MessageButtons, MessageDialog, MessageDialogResult, MessageLevel};
use serde_json::json;

const BUKVARIX_URL: &str = "https://api.bukvarix.com/v1/mkeywords/";
const PROJECT_EXT: &str = "el";
const NO_GROUP_KEY: &str = "__NO_GROUP__";
const DEFAULT_COLS: [&str; 4] = ["Фраза", "Частотность", "!Частостность", "[!Частостность]"];

type Table = (Vec<String>, Vec<Vec<String>>);

// ===========================
//  ДИАЛОГИ
// ===========================
fn info(title: &str, msg: &str) {
    MessageDialog::new().set_level(MessageLevel::Info).set_title(title).set_description(msg).show();
}
fn error(msg: &str) {
    MessageDialog::new().set_level(MessageLevel::Error).set_title("Ошибка").set_description(msg).show();
}
fn confirm(title: &str, msg: &str) -> bool {
    MessageDialog::new()
        .set_title(title)
        .set_description(msg)
        .set_buttons(MessageButtons::YesNo)
        .show()
        == MessageDialogResult::Yes
}

// ===========================
//  МОДЕЛЬ ДАННЫХ
// ===========================
#[derive(Clone, Default)]
struct Row {
    group: String,
    cells: Vec<String>, // cells[0] — фраза
}
impl Row {
    fn phrase(&self) -> &str {
        self.cells.first().map(String::as_str).unwrap_or("")
    }
}

#[derive(Clone, PartialEq)]
enum GroupFilter {
    All,
    NoGroup,
    Named(String),
}

#[derive(Clone, Copy)]
enum SortCol {
    Group,
    Cell(usize),
}

// Only presentation state. Group names, row format and business logic stay unchanged.
#[derive(Clone, Copy, PartialEq, Eq)]
enum WorkspaceTab {
    Files,
    Collect,
    Clean,
    Groups,
    Smart,
    Export,
}

#[derive(Default)]
struct GroupNode {
    label: String,
    full_name: String,
    is_group: bool,
    count: usize,
    children: BTreeMap<String, GroupNode>,
}

fn build_group_tree(groups: &[String], rows: &[Row]) -> BTreeMap<String, GroupNode> {
    let mut roots: BTreeMap<String, GroupNode> = BTreeMap::new();
    for name in groups {
        let parts: Vec<&str> = name.split('/').collect();
        let mut children = &mut roots;
        let mut path = String::new();
        for (i, part) in parts.iter().enumerate() {
            if i != 0 { path.push('/'); }
            path.push_str(part);
            let node = children.entry((*part).to_string()).or_insert_with(|| GroupNode {
                label: part.trim().to_string(),
                full_name: path.clone(),
                ..Default::default()
            });
            if i + 1 == parts.len() { node.is_group = true; }
            children = &mut node.children;
        }
    }
    let mut counts: HashMap<&str, usize> = HashMap::new();
    for r in rows {
        if !r.group.is_empty() { *counts.entry(r.group.as_str()).or_insert(0) += 1; }
    }
    fn calc(node: &mut GroupNode, counts: &HashMap<&str, usize>) -> usize {
        let own = counts.get(node.full_name.as_str()).copied().unwrap_or(0);
        let nested: usize = node.children.values_mut().map(|n| calc(n, counts)).sum();
        node.count = own + nested;
        node.count
    }
    for node in roots.values_mut() { calc(node, &counts); }
    roots
}

fn group_matches(node: &GroupNode, query: &str) -> bool {
    query.is_empty()
        || node.full_name.to_lowercase().contains(query)
        || node.children.values().any(|n| group_matches(n, query))
}

fn draw_group_node(
    ui: &mut egui::Ui,
    node: &GroupNode,
    current: &GroupFilter,
    query: &str,
    selected: &mut Option<GroupFilter>,
    drop_target: &mut Option<(GroupFilter, Vec<String>)>,
) {
    if !group_matches(node, query) { return; }
    let title = format!("{}   {}", node.label, node.count);
    if node.children.is_empty() {
        let target = GroupFilter::Named(node.full_name.clone());
        let response = ui.selectable_label(*current == target, title.clone());
        if response.clicked() { *selected = Some(target.clone()); }
        if let Some(payload) = response.dnd_release_payload::<DragPayload>() {
            *drop_target = Some((target, payload.0.clone()));
        }
    } else {
        let response = egui::CollapsingHeader::new(title.clone())
            .id_salt(("tree_node", &node.full_name))
            .default_open(true)
            .show(ui, |ui| {
                if node.is_group {
                    let target = GroupFilter::Named(node.full_name.clone());
                    let pick = ui.selectable_label(*current == target, "Все ключи этой ветки");
                    if pick.clicked() { *selected = Some(target.clone()); }
                    if let Some(payload) = pick.dnd_release_payload::<DragPayload>() {
                        *drop_target = Some((target, payload.0.clone()));
                    }
                }
                for child in node.children.values() {
                    draw_group_node(ui, child, current, query, selected, drop_target);
                }
            });
        if node.is_group {
            if let Some(payload) = response.header_response.dnd_release_payload::<DragPayload>() {
                *drop_target = Some((GroupFilter::Named(node.full_name.clone()), payload.0.clone()));
            }
        }
    }
}

// Vector-only navigation icons: do not rely on font glyph availability.
// This helper draws decorative UI shapes; it does not modify project data or actions.
fn paint_nav_icon(painter: &egui::Painter, origin: egui::Pos2, tab: WorkspaceTab, color: egui::Color32) {
    let point = |x: f32, y: f32| egui::pos2(origin.x + x, origin.y + y);
    let stroke = egui::Stroke::new(1.4, color);
    match tab {
        WorkspaceTab::Files => {
            painter.rect_stroke(egui::Rect::from_min_max(point(2.0, 0.5), point(12.5, 14.5)), 1.0, stroke);
            for y in [5.0, 8.0, 11.0] {
                painter.line_segment([point(4.5, y), point(10.0, y)], stroke);
            }
        }
        WorkspaceTab::Collect => {
            painter.line_segment([point(7.0, 1.0), point(7.0, 10.0)], stroke);
            painter.line_segment([point(3.5, 7.0), point(7.0, 10.5)], stroke);
            painter.line_segment([point(10.5, 7.0), point(7.0, 10.5)], stroke);
            painter.line_segment([point(2.5, 12.0), point(11.5, 12.0)], stroke);
        }
        WorkspaceTab::Clean => {
            painter.line_segment([point(7.0, 1.0), point(7.0, 13.0)], stroke);
            painter.line_segment([point(1.0, 7.0), point(13.0, 7.0)], stroke);
            painter.line_segment([point(3.0, 3.0), point(11.0, 11.0)], stroke);
            painter.line_segment([point(3.0, 11.0), point(11.0, 3.0)], stroke);
        }
        WorkspaceTab::Groups => {
            painter.rect_stroke(egui::Rect::from_min_max(point(1.0, 3.0), point(8.0, 11.0)), 1.0, stroke);
            painter.rect_stroke(egui::Rect::from_min_max(point(5.5, 5.5), point(13.0, 13.5)), 1.0, stroke);
        }
        WorkspaceTab::Smart => {
            painter.circle_stroke(point(6.0, 6.0), 4.0, stroke);
            painter.line_segment([point(9.0, 9.0), point(13.0, 13.0)], stroke);
            painter.line_segment([point(6.0, 4.0), point(6.0, 8.0)], stroke);
            painter.line_segment([point(4.0, 6.0), point(8.0, 6.0)], stroke);
        }
        WorkspaceTab::Export => {
            painter.rect_stroke(egui::Rect::from_min_max(point(1.0, 6.0), point(9.5, 13.0)), 1.0, stroke);
            painter.line_segment([point(5.0, 10.0), point(13.0, 2.0)], stroke);
            painter.line_segment([point(8.0, 2.0), point(13.0, 2.0)], stroke);
            painter.line_segment([point(13.0, 2.0), point(13.0, 7.0)], stroke);
        }
    }
}

// ===========================
//  АНАЛИЗ СЛОВ (НЕ МЕНЯЕТ ВЕТКИ И ЧАСТОТНОСТИ)
// ===========================
struct SmartGroup {
    key: String,
    forms: Vec<String>,
    rows: Vec<usize>, // индексы строк исходного набора, каждая строка не более одного раза
}

struct SmartState {
    groups: Vec<SmartGroup>,
    selected: HashSet<String>,
    focused: Option<String>,
    query: String,
    min_count: usize,
    stems: bool,
    all_rows: bool,
    dirty: bool,
    feedback: String,
}

impl Default for SmartState {
    fn default() -> Self {
        Self {
            groups: Vec::new(),
            selected: HashSet::new(),
            focused: None,
            query: String::new(),
            min_count: 2,
            stems: true,
            all_rows: false,
            dirty: true,
            feedback: String::new(),
        }
    }
}

/// Группы по словам; при stems=true русские/английские словоформы
/// объединяются по алгоритму Snowball. Сопоставление с исходными
/// фразами основано на целых токенах, а НЕ на подстроке.
fn build_smart_groups(rows: &[Row], row_ids: &[usize], stems: bool) -> Vec<SmartGroup> {
    let pattern = Regex::new(r"[\p{L}\p{N}]+").expect("правильное выражение для слов");
    let ru = Stemmer::create(Algorithm::Russian);
    let en = Stemmer::create(Algorithm::English);
    let mut groups: HashMap<String, (BTreeSet<String>, Vec<usize>)> = HashMap::new();

    for &row_id in row_ids {
        let Some(row) = rows.get(row_id) else { continue; };
        let lower = row.phrase().to_lowercase().replace('ё', "е");
        let mut unique_keys = HashSet::new();
        for found in pattern.find_iter(&lower) {
            let form = found.as_str();
            let key = if !stems {
                format!("word:{form}")
            } else if form.chars().all(|c| ('а'..='я').contains(&c)) {
                format!("ru:{}", ru.stem(form))
            } else if form.bytes().all(|c| c.is_ascii_alphabetic()) {
                format!("en:{}", en.stem(form))
            } else {
                format!("word:{form}")
            };
            let entry = groups.entry(key.clone()).or_default();
            entry.0.insert(form.to_string());
            if unique_keys.insert(key) { entry.1.push(row_id); }
        }
    }

    let mut result: Vec<SmartGroup> = groups.into_iter().map(|(key, (forms, rows))| {
        SmartGroup { key, forms: forms.into_iter().collect(), rows }
    }).collect();
    result.sort_unstable_by(|a, b| {
        b.rows.len().cmp(&a.rows.len())
            .then_with(|| a.forms.first().cmp(&b.forms.first()))
    });
    result
}

struct DragPayload(Vec<String>);

struct Picker {
    phrases: Vec<String>,
    sel: usize, // 0 = "Без ветки"
}

struct BukRow {
    phrase: String,
    broad: String,
    exact: String,
}

#[derive(Default)]
struct BukState {
    open: bool,
    input: String,
    api_key: String,
    num: u32,
    busy: bool,
    rx: Option<Receiver<Result<Vec<BukRow>, String>>>,
}

fn parse_num(s: &str) -> Option<f64> {
    s.trim().parse::<f64>().ok().filter(|v| !v.is_nan())
}

// ===========================
//  ЗАГРУЗКА ТАБЛИЦ
// ===========================
fn normalize_headers(h: Vec<String>) -> Vec<String> {
    let mut seen: HashMap<String, usize> = HashMap::new();
    h.into_iter()
        .enumerate()
        .map(|(i, name)| {
            let name = if name.trim().is_empty() { format!("Unnamed: {i}") } else { name };
            let n = seen.entry(name.clone()).or_insert(0);
            let out = if *n == 0 { name.clone() } else { format!("{name}.{n}") };
            *n += 1;
            out
        })
        .collect()
}

fn read_csv_text(text: &str, delim: u8) -> Result<Table, Box<dyn Error>> {
    let text = text.trim_start_matches('\u{feff}');
    let mut rdr = csv::ReaderBuilder::new()
        .delimiter(delim)
        .has_headers(true)
        .flexible(true)
        .from_reader(text.as_bytes());
    let headers = normalize_headers(rdr.headers()?.iter().map(String::from).collect());
    let n = headers.len();
    let mut rows = Vec::new();
    for rec in rdr.records() {
        let rec = rec?;
        let mut v: Vec<String> = rec.iter().map(String::from).collect();
        v.resize(n, String::new());
        rows.push(v);
    }
    Ok((headers, rows))
}

fn cell_str(c: &Data) -> String {
    match c {
        Data::Empty => String::new(),
        o => o.to_string(),
    }
}

fn read_table(path: &Path) -> Result<Table, Box<dyn Error>> {
    let ext = path.extension().and_then(|e| e.to_str()).unwrap_or("").to_lowercase();
    if ext == "csv" {
        let text = String::from_utf8(std::fs::read(path)?)?;
        let first = text.trim_start_matches('\u{feff}').lines().next().unwrap_or("");
        let delim = if first.contains(';') { b';' } else { b',' };
        read_csv_text(&text, delim)
    } else {
        let mut wb = open_workbook_auto(path)?;
        let name = wb.sheet_names().first().cloned().ok_or("В книге нет листов")?;
        let range = wb.worksheet_range(&name)?;
        let mut it = range.rows();
        let headers = normalize_headers(
            it.next().map(|r| r.iter().map(cell_str).collect()).unwrap_or_default(),
        );
        let n = headers.len();
        let rows = it
            .map(|r| {
                let mut v: Vec<String> = r.iter().map(cell_str).collect();
                v.resize(n, String::new());
                v
            })
            .collect();
        Ok((headers, rows))
    }
}

fn clean_group(s: &str) -> String {
    let t = s.trim();
    if t == "nan" || t == "None" { String::new() } else { t.to_string() }
}

/// Выделяет служебную колонку ветки ("_group", иначе "Группа") в Row.group.
fn build_rows(headers: Vec<String>, rows: Vec<Vec<String>>) -> (Vec<String>, Vec<Row>) {
    let gidx = headers
        .iter()
        .position(|h| h == "_group")
        .or_else(|| headers.iter().position(|h| h == "Группа"));
    let mut cols = headers;
    if let Some(g) = gidx {
        cols.remove(g);
    }
    let out = rows
        .into_iter()
        .map(|mut v| {
            let group = match gidx {
                Some(g) => clean_group(&v.remove(g)),
                None => String::new(),
            };
            Row { group, cells: v }
        })
        .collect();
    (cols, out)
}

// ===========================
//  НОРМАЛИЗАЦИЯ ДЛЯ ДЕДУПЛИКАЦИИ
// ===========================
fn norm_exact(v: &str) -> String {
    let s = v.to_lowercase().replace('ё', "е");
    s.split_whitespace().collect::<Vec<_>>().join(" ")
}

fn norm_tokensort(v: &str) -> String {
    let s = v.to_lowercase().replace('ё', "е");
    let re = Regex::new(r"[0-9a-zа-я]+").unwrap();
    let mut t: Vec<&str> = re.find_iter(&s).map(|m| m.as_str()).collect();
    t.sort();
    t.join(" ")
}

// ===========================
//  XLSX
// ===========================
fn sheet_name(raw: &str) -> String {
    let mut name: String = raw
        .trim()
        .chars()
        .map(|c| if "[]:*?/\\".contains(c) { '_' } else { c })
        .collect();
    name = name.trim().to_string();
    if name.is_empty() {
        name = "Без ветки".into();
    }
    name.chars().take(31).collect()
}

fn unique_sheet_name(base: String, used: &mut HashSet<String>) -> String {
    if used.insert(base.clone()) {
        return base;
    }
    let mut i = 2;
    loop {
        let suffix = format!("_{i}");
        let keep = 31 - suffix.chars().count();
        let cand: String = base.chars().take(keep).collect::<String>() + &suffix;
        if used.insert(cand.clone()) {
            return cand;
        }
        i += 1;
    }
}

// ===========================
//  BUKVARIX
// ===========================
/// Печатает всю цепочку причин: reqwest -> транспорт -> TLS/proxy/IO.
fn reqwest_error_details(action: &str, e: &reqwest::Error) -> String {
    let mut details = format!("{action}: {e}");
    let mut cause = e.source();
    let mut depth = 0;
    while let Some(err) = cause {
        details.push_str(&format!("\nПричина {}: {err}", depth + 1));
        cause = err.source();
        depth += 1;
        if depth >= 12 { break; }
    }
    details
}

fn bukvarix_request(seeds: &[String], api_key: &str, num: u32) -> Result<Vec<BukRow>, String> {
    let q = seeds.join("\r\n");
    let num_s = num.to_string();
    let key = if api_key.trim().is_empty() { "free" } else { api_key.trim() };

    // На Windows/macOS используем системный TLS. На Linux — rustls.
    // Сетевая операция по-прежнему выполняется в отдельном потоке от egui.
    let builder = reqwest::blocking::Client::builder()
        .timeout(Duration::from_secs(45));

    #[cfg(any(target_os = "windows", target_os = "macos"))]
    let builder = builder.tls_backend_native();

    #[cfg(not(any(target_os = "windows", target_os = "macos")))]
    let builder = builder.tls_backend_rustls();

    let client = builder.build()
        .map_err(|e| reqwest_error_details("Не удалось создать HTTP-клиент Bukvarix", &e))?;

    let response = client
        .post(BUKVARIX_URL)
        .form(&[
            ("api_key", key),
            ("q", q.as_str()),
            ("format", "csv"),
            ("header", "0"),
            ("num", num_s.as_str()),
            ("bom", "1"),
        ])
        .send()
        .map_err(|e| reqwest_error_details("Не удалось выполнить запрос к Bukvarix", &e))?;

    let status = response.status();
    let text = response
        .text()
        .map_err(|e| reqwest_error_details("Не удалось прочитать ответ Bukvarix", &e))?;
    if !status.is_success() {
        return Err(format!("Bukvarix вернул HTTP {status}: {text}"));
    }
    let text = text.trim_start_matches('\u{feff}');
    let mut rdr = csv::ReaderBuilder::new()
        .delimiter(b';')
        .has_headers(false)
        .flexible(true)
        .from_reader(text.as_bytes());
    let mut out = Vec::new();
    for rec in rdr.records() {
        let rec = rec.map_err(|e| e.to_string())?;
        let num_or_empty = |i: usize| {
            let s = rec.get(i).unwrap_or("").trim();
            if parse_num(s).is_some() { s.to_string() } else { String::new() }
        };
        out.push(BukRow {
            phrase: rec.get(0).unwrap_or("").to_string(),
            broad: num_or_empty(3),
            exact: num_or_empty(4),
        });
    }
    Ok(out)
}

// ===========================
//  ПРИЛОЖЕНИЕ
// ===========================
struct App {
    columns: Vec<String>,
    rows: Vec<Row>,
    has_data: bool,

    stop_words: BTreeSet<String>,
    stop_re: Option<Regex>,
    history: Vec<BTreeSet<String>>,

    search: String,
    group_filter: GroupFilter,
    groups: Vec<String>,
    last_target_group: Option<String>,
    project_path: Option<PathBuf>,
    sort_state: HashMap<usize, bool>,

    view: Vec<usize>,
    selected: HashSet<String>,
    last_clicked: Option<usize>,

    show_stop_words: bool,
    word_dialog: Option<Vec<String>>,
    new_group: Option<String>,
    picker: Option<Picker>,
    show_contact: bool,
    buk: BukState,

    // UI-only settings: no changes to project file format.
    tab: WorkspaceTab,
    group_search: String,
    visible_columns: Vec<bool>,
    show_details: bool,
    smart: SmartState,
}

impl App {
    fn new() -> Self {
        Self {
            columns: DEFAULT_COLS.iter().map(|s| s.to_string()).collect(),
            rows: vec![],
            has_data: false,
            stop_words: BTreeSet::new(),
            stop_re: None,
            history: vec![],
            search: String::new(),
            group_filter: GroupFilter::All,
            groups: vec![],
            last_target_group: None,
            project_path: None,
            sort_state: HashMap::new(),
            view: vec![],
            selected: HashSet::new(),
            last_clicked: None,
            show_stop_words: false,
            word_dialog: None,
            new_group: None,
            picker: None,
            show_contact: false,
            buk: BukState { api_key: "free".into(), num: 250, ..Default::default() },
            tab: WorkspaceTab::Files,
            group_search: String::new(),
            visible_columns: vec![true; DEFAULT_COLS.len()],
            show_details: false,
            smart: SmartState::default(),
        }
    }

    // ---------- данные ----------
    fn set_data(&mut self, headers: Vec<String>, raw: Vec<Vec<String>>) {
        let (cols, rows) = build_rows(headers, raw);
        self.columns = if cols.is_empty() {
            DEFAULT_COLS.iter().map(|s| s.to_string()).collect()
        } else {
            cols
        };
        let n = self.columns.len();
        self.visible_columns = vec![true; n];
        self.rows = rows;
        for r in &mut self.rows {
            r.cells.resize(n, String::new());
        }
        self.has_data = true;
        self.search.clear();
        self.group_filter = GroupFilter::All;
        self.refresh_all();
    }

    fn refresh_all(&mut self) {
        self.rebuild_groups();
        self.recompute_view();
    }

    fn rebuild_stop_re(&mut self) {
        self.stop_re = if self.stop_words.is_empty() {
            None
        } else {
            let alts: Vec<String> = self.stop_words.iter().map(|w| regex::escape(w)).collect();
            Regex::new(&format!(r"(?i)\b({})\b", alts.join("|"))).ok()
        };
    }

    fn passes_stop(&self, r: &Row) -> bool {
        match &self.stop_re {
            Some(re) => !re.is_match(r.phrase()),
            None => true,
        }
    }

    fn rebuild_groups(&mut self) {
        let mut set: BTreeSet<String> = self
            .groups
            .iter()
            .filter(|g| !g.trim().is_empty() && g.to_lowercase() != "nan")
            .cloned()
            .collect();
        for r in &self.rows {
            let g = r.group.trim();
            if !g.is_empty() && g.to_lowercase() != "nan" {
                set.insert(g.to_string());
            }
        }
        self.groups = set.into_iter().collect();
    }

    fn recompute_view(&mut self) {
        let words: Vec<String> =
            self.search.to_lowercase().split_whitespace().map(String::from).collect();
        let view: Vec<usize> = self
            .rows
            .iter()
            .enumerate()
            .filter(|(_, r)| {
                if !self.passes_stop(r) {
                    return false;
                }
                let g_ok = match &self.group_filter {
                    GroupFilter::All => true,
                    GroupFilter::NoGroup => r.group.trim().is_empty(),
                    GroupFilter::Named(g) => &r.group == g,
                };
                if !g_ok {
                    return false;
                }
                if words.is_empty() {
                    return true;
                }
                let p = r.phrase().to_lowercase();
                words.iter().all(|w| p.contains(w.as_str()))
            })
            .map(|(i, _)| i)
            .collect();
        self.view = view;
        self.smart.dirty = true;
        self.selected.clear();
        self.last_clicked = None;
    }

    // ---------- загрузка / сохранение ----------
    fn load_data(&mut self) {
        let Some(path) = FileDialog::new()
            .add_filter("CSV и Excel", &["csv", "xlsx"])
            .add_filter("CSV files", &["csv"])
            .add_filter("Excel files", &["xlsx"])
            .pick_file()
        else {
            return;
        };
        match read_table(&path) {
            Ok((h, r)) => self.set_data(h, r),
            Err(e) => error(&format!("Не удалось загрузить файл:\n{e}")),
        }
    }

    fn save_file(&self) {
        let rows: Vec<&Row> = self.rows.iter().filter(|r| self.passes_stop(r)).collect();
        if !self.has_data || rows.is_empty() {
            info("Информация", "Нет данных для сохранения.");
            return;
        }
        let Some(mut path) = FileDialog::new()
            .add_filter("Excel files", &["xlsx"])
            .add_filter("CSV files", &["csv"])
            .set_file_name("export.xlsx")
            .save_file()
        else {
            return;
        };
        let ext = path.extension().and_then(|e| e.to_str()).map(|e| e.to_lowercase());
        if ext.is_none() {
            path.set_extension("xlsx");
        }
        if ext.as_deref() == Some("csv") {
            let res = (|| -> Result<(), Box<dyn Error>> {
                let mut w = csv::Writer::from_path(&path)?;
                let mut h = vec!["Группа".to_string()];
                h.extend(self.columns.iter().cloned());
                w.write_record(&h)?;
                for r in &rows {
                    let mut rec = vec![r.group.clone()];
                    rec.extend(r.cells.iter().cloned());
                    w.write_record(&rec)?;
                }
                w.flush()?;
                Ok(())
            })();
            match res {
                Ok(_) => info("Успех", "CSV успешно сохранён!"),
                Err(e) => error(&format!("Не удалось сохранить файл:\n{e}")),
            }
        } else {
            match self.save_xlsx(&path, &rows) {
                Ok(n) => info("Успех", &format!("Excel сохранён.\nЛистов: {n} (ветки + стоп-слова).")),
                Err(e) => error(&format!("Не удалось сохранить файл:\n{e}")),
            }
        }
    }

    fn save_xlsx(&self, path: &Path, rows: &[&Row]) -> Result<usize, Box<dyn Error>> {
        use rust_xlsxwriter::Workbook;
        let mut wb = Workbook::new();
        let mut used = HashSet::new();
        // BTreeSet: пустая ветка "" автоматически первая
        let groups: BTreeSet<String> = rows.iter().map(|r| r.group.trim().to_string()).collect();

        for g in &groups {
            let name = unique_sheet_name(sheet_name(g), &mut used);
            let ws = wb.add_worksheet();
            ws.set_name(&name)?;
            for (c, h) in self.columns.iter().enumerate() {
                ws.write_string(0, c as u16, h)?;
            }
            let mut rn = 1u32;
            for r in rows.iter().filter(|r| r.group.trim() == g) {
                for (c, v) in r.cells.iter().enumerate() {
                    if v.is_empty() {
                        continue;
                    }
                    match (c > 0, parse_num(v)) {
                        (true, Some(n)) if n.is_finite() => ws.write_number(rn, c as u16, n)?,
                        _ => ws.write_string(rn, c as u16, v)?,
                    };
                }
                rn += 1;
            }
        }

        let name = unique_sheet_name(sheet_name("Стоп-слова"), &mut used);
        let ws = wb.add_worksheet();
        ws.set_name(&name)?;
        ws.write_string(0, 0, "Стоп-слова")?;
        let mut sw: Vec<&String> = self.stop_words.iter().collect();
        sw.sort_by_key(|w| w.to_lowercase());
        for (i, w) in sw.iter().enumerate() {
            ws.write_string(i as u32 + 1, 0, w.as_str())?;
        }
        wb.save(path)?;
        Ok(groups.len() + 1)
    }

    // ---------- проекты ----------
    fn save_project_as(&mut self) {
        let Some(mut path) = FileDialog::new()
            .add_filter(format!("Проект (.{PROJECT_EXT})"), &[PROJECT_EXT])
            .set_file_name(format!("project.{PROJECT_EXT}"))
            .save_file()
        else {
            return;
        };
        if path.extension().and_then(|e| e.to_str()).map(|e| e.to_lowercase()).as_deref()
            != Some(PROJECT_EXT)
        {
            path.set_extension(PROJECT_EXT);
        }
        self.project_path = Some(path.clone());
        self.save_project_to_path(&path);
    }

    fn save_project(&mut self) {
        match self.project_path.clone() {
            Some(p) => self.save_project_to_path(&p),
            None => self.save_project_as(),
        }
    }

    fn save_project_to_path(&self, path: &Path) {
        if !self.has_data {
            info("Информация", "Нет данных для сохранения проекта. Сначала загрузите файл/парсинг.");
            return;
        }
        let res = (|| -> Result<(), Box<dyn Error>> {
            let mut cols = self.columns.clone();
            cols.push("_group".into());
            let meta = json!({
                "version": 1,
                "groups": self.groups,
                "stop_words": self.stop_words.iter().collect::<Vec<_>>(),
                "current_search_query": self.search,
                "current_group_filter": match &self.group_filter {
                    GroupFilter::All => None,
                    GroupFilter::NoGroup => Some(NO_GROUP_KEY.to_string()),
                    GroupFilter::Named(g) => Some(g.clone()),
                },
                "last_target_group": self.last_target_group,
                "columns": cols,
            });
            let mut w = csv::Writer::from_writer(vec![]);
            w.write_record(&cols)?;
            for r in &self.rows {
                let mut rec = r.cells.clone();
                rec.push(r.group.clone());
                w.write_record(&rec)?;
            }
            let data = w.into_inner().map_err(|e| e.to_string())?;
            let mut z = zip::ZipWriter::new(File::create(path)?);
            let opt = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated);
            z.start_file("data.csv", opt)?;
            z.write_all(&data)?;
            z.start_file("meta.json", opt)?;
            z.write_all(serde_json::to_string_pretty(&meta)?.as_bytes())?;
            z.finish()?;
            Ok(())
        })();
        match res {
            Ok(_) => info("Успех", &format!("Проект сохранён:\n{}", path.display())),
            Err(e) => error(&format!("Не удалось сохранить проект:\n{e}")),
        }
    }

    fn open_project(&mut self) {
        let Some(path) = FileDialog::new()
            .add_filter(format!("Проект (.{PROJECT_EXT})"), &[PROJECT_EXT])
            .pick_file()
        else {
            return;
        };
        let res = (|| -> Result<(Table, serde_json::Value), Box<dyn Error>> {
            let mut z = zip::ZipArchive::new(File::open(&path)?)?;
            let bad = "Неверный формат проекта: нет data.csv или meta.json";
            let mut data = String::new();
            z.by_name("data.csv").map_err(|_| bad)?.read_to_string(&mut data)?;
            let mut m = String::new();
            z.by_name("meta.json").map_err(|_| bad)?.read_to_string(&mut m)?;
            Ok((read_csv_text(&data, b',')?, serde_json::from_str(&m)?))
        })();
        match res {
            Ok(((h, r), meta)) => {
                self.set_data(h, r);
                let strs = |k: &str| -> Vec<String> {
                    meta[k]
                        .as_array()
                        .map(|a| a.iter().filter_map(|v| v.as_str().map(String::from)).collect())
                        .unwrap_or_default()
                };
                self.groups = strs("groups");
                self.stop_words = strs("stop_words").into_iter().collect();
                self.history.clear();
                self.search = meta["current_search_query"].as_str().unwrap_or("").to_string();
                self.group_filter = match meta["current_group_filter"].as_str() {
                    None => GroupFilter::All,
                    Some(NO_GROUP_KEY) => GroupFilter::NoGroup,
                    Some(g) => GroupFilter::Named(g.to_string()),
                };
                self.last_target_group = meta["last_target_group"].as_str().map(String::from);
                self.rebuild_stop_re();
                self.refresh_all();
                info("Успех", &format!("Проект открыт:\n{}", path.display()));
                self.project_path = Some(path);
            }
            Err(e) => error(&format!("Не удалось открыть проект:\n{e}")),
        }
    }

    // ---------- стоп-слова ----------
    fn add_stop_word(&mut self, word: String) {
        if self.stop_words.contains(&word) {
            info("Информация", &format!("Слово '{word}' уже в списке стоп-слов."));
            return;
        }
        self.history.push(self.stop_words.clone());
        self.stop_words.insert(word);
        self.rebuild_stop_re();
        self.recompute_view();
    }

    fn remove_stop_word(&mut self, word: &str) {
        if self.stop_words.contains(word) {
            self.history.push(self.stop_words.clone());
            self.stop_words.remove(word);
            self.rebuild_stop_re();
            self.recompute_view();
        }
    }

    fn undo_last_action(&mut self) {
        match self.history.pop() {
            Some(prev) => {
                self.stop_words = prev;
                self.rebuild_stop_re();
                self.recompute_view();
            }
            None => info("Отмена", "Нет действий для отмены."),
        }
    }

    fn sorted_stop_words(&self) -> Vec<String> {
        let mut v: Vec<String> = self.stop_words.iter().cloned().collect();
        v.sort_by_key(|w| w.to_lowercase());
        v
    }

    fn save_stop_words_to_file(&self) {
        if self.stop_words.is_empty() {
            info("Информация", "Список стоп-слов пуст.");
            return;
        }
        let Some(path) = FileDialog::new().add_filter("Text files", &["txt"]).set_file_name("stopwords.txt").save_file()
        else {
            return;
        };
        match std::fs::write(&path, self.sorted_stop_words().join("\n")) {
            Ok(_) => info("Успех", "Список стоп-слов сохранен!"),
            Err(e) => error(&format!("Не удалось сохранить файл:\n{e}")),
        }
    }

    fn load_stop_words_from_file(&mut self) {
        let Some(path) = FileDialog::new().add_filter("Text files", &["txt"]).pick_file() else {
            return;
        };
        match std::fs::read_to_string(&path) {
            Ok(text) => {
                let words: Vec<String> = text
                    .trim_start_matches('\u{feff}')
                    .lines()
                    .map(|l| l.trim().to_string())
                    .filter(|l| !l.is_empty())
                    .collect();
                if words.is_empty() {
                    info("Информация", "Файл стоп-слов пуст.");
                } else {
                    self.history.push(self.stop_words.clone());
                    self.stop_words.extend(words);
                    self.rebuild_stop_re();
                    self.recompute_view();
                    info("Успех", "Стоп-слова загружены и применены.");
                }
            }
            Err(e) => error(&format!("Не удалось загрузить файл:\n{e}")),
        }
    }

    fn open_word_selection(&mut self, phrase: &str) {
        let strip: &[char] = &[',', '.', '!', '?', ';', ':', '"', '\'', '(', ')', '[', ']', '{', '}'];
        let set: BTreeSet<String> = phrase
            .split_whitespace()
            .map(|w| w.trim_matches(strip).to_string())
            .filter(|w| !w.is_empty())
            .collect();
        if set.is_empty() {
            info("Информация", "Не удалось выделить слова из фразы.");
            return;
        }
        let mut v: Vec<String> = set.into_iter().collect();
        v.sort_by_key(|w| w.to_lowercase());
        self.word_dialog = Some(v);
    }

    // ---------- сортировка ----------
    fn sort_by(&mut self, col: SortCol, asc: bool) {
        if !self.has_data || self.rows.is_empty() {
            return;
        }
        let dir = move |o: Ordering| if asc { o } else { o.reverse() };
        match col {
            SortCol::Group => self.rows.sort_by(|a, b| dir(a.group.cmp(&b.group))),
            SortCol::Cell(i) => {
                if i >= self.columns.len() {
                    return;
                }
                let get = |r: &Row| r.cells.get(i).cloned().unwrap_or_default();
                let numeric = self.rows.iter().any(|r| parse_num(&get(r)).is_some());
                if numeric {
                    let mut keyed: Vec<(Option<f64>, Row)> = std::mem::take(&mut self.rows)
                        .into_iter()
                        .map(|r| (parse_num(&get(&r)), r))
                        .collect();
                    keyed.sort_by(|a, b| match (a.0, b.0) {
                        (Some(x), Some(y)) => dir(x.partial_cmp(&y).unwrap_or(Ordering::Equal)),
                        (Some(_), None) => Ordering::Less, // NaN всегда в конце
                        (None, Some(_)) => Ordering::Greater,
                        (None, None) => Ordering::Equal,
                    });
                    self.rows = keyed.into_iter().map(|(_, r)| r).collect();
                } else {
                    self.rows.sort_by(|a, b| dir(get(a).cmp(&get(b))));
                }
            }
        }
        self.recompute_view();
    }

    fn sort_by_tree_column(&mut self, display_idx: usize) {
        let asc = self.sort_state.get(&display_idx).copied().unwrap_or(true);
        self.sort_state.insert(display_idx, !asc);
        let col = if display_idx == 0 { SortCol::Group } else { SortCol::Cell(display_idx - 1) };
        self.sort_by(col, asc);
    }

    // ---------- дедупликация ----------
    fn dedupe(&mut self, key_fn: fn(&str) -> String, title: &str, rule: &str) {
        if !self.has_data || self.rows.is_empty() {
            info("Информация", "Нет данных для дедупликации.");
            return;
        }
        let f1 = self
            .columns
            .iter()
            .position(|c| c == "Частотность")
            .or(if self.columns.len() >= 2 { Some(1) } else { None });
        let f2 = if self.columns.len() >= 3 { Some(2) } else { None };
        let num = |r: &Row, i: Option<usize>| -> f64 {
            i.and_then(|i| r.cells.get(i)).and_then(|s| parse_num(s)).unwrap_or(-1.0)
        };
        let better = |a: &Row, b: &Row| -> bool {
            let (a1, b1) = (num(a, f1), num(b, f1));
            if a1 != b1 {
                return a1 > b1;
            }
            let (a2, b2) = (num(a, f2), num(b, f2));
            if a2 != b2 {
                return a2 > b2;
            }
            a.phrase().chars().count() < b.phrase().chars().count()
        };

        let mut best: HashMap<String, usize> = HashMap::new();
        for (i, r) in self.rows.iter().enumerate() {
            if !self.passes_stop(r) {
                continue;
            }
            let k = key_fn(r.phrase());
            let prev = best.get(&k).copied();
            match prev {
                None => {
                    best.insert(k, i);
                }
                Some(j) => {
                    if better(r, &self.rows[j]) {
                        best.insert(k, i);
                    }
                }
            }
        }
        let keep: HashSet<usize> = best.values().copied().collect();
        let stop = self.stop_re.clone();
        let before = self.rows.len();
        let mut idx = 0usize;
        self.rows.retain(|r| {
            let i = idx;
            idx += 1;
            let passes = stop.as_ref().map_or(true, |re| !re.is_match(r.phrase()));
            !passes || keep.contains(&i)
        });
        let removed = before - self.rows.len();
        self.refresh_all();
        info(
            "Готово",
            &format!("{title}: {removed}\nОсталось строк: {}\n\nПравило: {rule}", self.view.len()),
        );
    }

    // ---------- ветки ----------
    fn assign_group(&mut self, phrases: &[String], target: &str) {
        let set: HashSet<&str> = phrases.iter().map(String::as_str).collect();
        for r in &mut self.rows {
            if set.contains(r.phrase()) {
                r.group = target.to_string();
            }
        }
        self.refresh_all();
    }

    fn open_picker(&mut self, phrases: Vec<String>) {
        if self.groups.is_empty() {
            info("Информация", "Сначала создайте ветку кнопкой '+ Ветка'.");
            return;
        }
        let sel = match &self.last_target_group {
            Some(g) if g.is_empty() => 0,
            Some(g) => self.groups.iter().position(|x| x == g).map(|i| i + 1).unwrap_or(1),
            None => 1,
        };
        self.picker = Some(Picker { phrases, sel });
    }

    fn selected_phrases(&self) -> Vec<String> {
        self.selected.iter().cloned().collect()
    }

    fn move_selected_to_group(&mut self) {
        if !self.has_data {
            info("Информация", "Сначала загрузите данные.");
            return;
        }
        let phrases = self.selected_phrases();
        if phrases.is_empty() {
            info("Информация", "Выделите строки в таблице.");
            return;
        }
        match self.group_filter.clone() {
            GroupFilter::NoGroup => self.assign_group(&phrases, ""),
            GroupFilter::All => self.open_picker(phrases),
            GroupFilter::Named(g) => self.assign_group(&phrases, &g),
        }
    }

    fn move_all_view_to_group(&mut self) {
        if !self.has_data || self.view.is_empty() {
            info("Информация", "Нет данных для переноса (таблица пустая).");
            return;
        }
        let target = match self.group_filter.clone() {
            GroupFilter::All => {
                info("Информация", "Выберите конкретную ветку справа (не 'Все').");
                return;
            }
            GroupFilter::NoGroup => String::new(),
            GroupFilter::Named(g) => g,
        };
        let phrases: Vec<String> =
            self.view.iter().map(|&i| self.rows[i].phrase().to_string()).collect();
        self.assign_group(&phrases, &target);
    }

    fn delete_group(&mut self) {
        let GroupFilter::Named(label) = self.group_filter.clone() else {
            info("Информация", "Эту системную ветку удалить нельзя.");
            return;
        };
        if !confirm(
            "Удалить ветку",
            &format!("Удалить ветку '{label}'?\n\nКлючи из этой ветки станут 'Без ветки'."),
        ) {
            return;
        }
        for r in &mut self.rows {
            if r.group == label {
                r.group.clear();
            }
        }
        self.groups.retain(|g| g != &label);
        self.group_filter = GroupFilter::All;
        self.refresh_all();
    }

    fn delete_selected_rows(&mut self) {
        if !self.has_data {
            return;
        }
        if self.selected.is_empty() {
            info("Информация", "Выделите строки в таблице для удаления.");
            return;
        }
        if !confirm(
            "Удаление",
            &format!(
                "Удалить выбранные строки: {}?\nЭто действие удалит фразы из набора данных.",
                self.selected.len()
            ),
        ) {
            return;
        }
        let set = std::mem::take(&mut self.selected);
        self.rows.retain(|r| !set.contains(r.phrase()));
        self.refresh_all();
    }

    // ---------- аналитическая группировка ----------
    fn rebuild_smart_groups(&mut self) {
        let row_ids: Vec<usize> = if self.smart.all_rows {
            (0..self.rows.len()).collect()
        } else {
            self.view.clone()
        };
        self.smart.groups = build_smart_groups(&self.rows, &row_ids, self.smart.stems);
        let keys: HashSet<&str> = self.smart.groups.iter().map(|g| g.key.as_str()).collect();
        self.smart.selected.retain(|key| keys.contains(key.as_str()));
        if self.smart.focused.as_ref().is_some_and(|key| !keys.contains(key.as_str())) {
            self.smart.focused = None;
        }
        if self.smart.focused.is_none() {
            self.smart.focused = self.smart.groups.first().map(|g| g.key.clone());
        }
        self.smart.dirty = false;
    }

    fn smart_selected_rows(&self) -> BTreeSet<usize> {
        self.smart.groups.iter()
            .filter(|g| self.smart.selected.contains(&g.key))
            .flat_map(|g| g.rows.iter().copied())
            .collect()
    }

    fn smart_add_stop_words(&mut self) {
        let forms: BTreeSet<String> = self.smart.groups.iter()
            .filter(|g| self.smart.selected.contains(&g.key))
            .flat_map(|g| g.forms.iter().cloned())
            .collect();
        if forms.is_empty() { return; }
        let before = self.stop_words.len();
        // Одна операция в истории для всех выбранных групп.
        self.history.push(self.stop_words.clone());
        self.stop_words.extend(forms);
        let added = self.stop_words.len() - before;
        self.smart.selected.clear();
        self.rebuild_stop_re();
        self.recompute_view();
        self.smart.feedback = format!(
            "Добавлено стоп-слов: {added}. Теперь они скрывают совпадающие фразы в таблице и при экспорте. «Назад» отменит операцию."
        );
    }

    fn smart_delete_phrases(&mut self) {
        let selected_rows = self.smart_selected_rows();
        if selected_rows.is_empty() { return; }
        if !confirm("Удалить совпадающие фразы", &format!(
            "Безвозвратно удалить {} строк из набора данных?\n\nСтроки совпадают хотя бы с одной выбранной группой слов. Перед массовым удалением рекомендуется сохранить проект.",
            selected_rows.len()
        )) { return; }
        let before = self.rows.len();
        let mut position = 0usize;
        self.rows.retain(|_| {
            let keep = !selected_rows.contains(&position);
            position += 1;
            keep
        });
        let deleted = before - self.rows.len();
        self.smart.selected.clear();
        self.refresh_all();
        self.smart.feedback = format!("Удалено строк: {deleted}. В отличие от стоп-слов, это физическое удаление из проекта.");
    }

    fn smart_group_panel(&mut self, ctx: &egui::Context) {
        if self.smart.dirty { self.rebuild_smart_groups(); }
        let mut reanalyze = false;
        let mut add_stop = false;
        let mut delete = false;
        let mut show_stop_words = false;
        let query = self.smart.query.trim().to_lowercase();
        let filtered: Vec<usize> = self.smart.groups.iter().enumerate()
            .filter(|(_, g)| g.rows.len() >= self.smart.min_count && (
                query.is_empty() || g.forms.iter().any(|f| f.contains(&query))
            ))
            .map(|(i, _)| i)
            .collect();
        let selected_count = self.smart.selected.len();
        let selected_rows = self.smart_selected_rows();
        let preview_rows: Vec<usize> = if selected_count != 0 {
            selected_rows.iter().copied().collect()
        } else {
            self.smart.focused.as_ref()
                .and_then(|key| self.smart.groups.iter().find(|g| &g.key == key))
                .map(|g| g.rows.clone())
                .unwrap_or_default()
        };
        let analysed_rows = if self.smart.all_rows { self.rows.len() } else { self.view.len() };

        egui::CentralPanel::default().show(ctx, |ui| {
            ui.heading("Анализ слов и массовая очистка");
            ui.add_space(7.0);
            ui.horizontal_wrapped(|ui| {
                if ui.checkbox(&mut self.smart.stems, "Объединять словоформы RU / EN").changed() {
                    reanalyze = true;
                }
                if ui.checkbox(&mut self.smart.all_rows, "Анализировать все, включая скрытые фразы").changed() {
                    reanalyze = true;
                }
                if ui.button("Обновить анализ").clicked() { reanalyze = true; }
            });
            ui.label(egui::RichText::new(
                "Словоформы объединяются приближённо: проверяй список совпадений перед удалением. Отключи опцию для строго одинаковых слов."
            ).small().color(egui::Color32::from_gray(142)));
            ui.add_space(8.0);
            ui.horizontal_wrapped(|ui| {
                ui.label(format!("Фраз изучено: {analysed_rows}   ·   Групп: {}", self.smart.groups.len()));
                ui.label(format!("Отмечено групп: {selected_count}   ·   Совпавших фраз: {}", selected_rows.len()));
            });
            ui.add_space(6.0);
            ui.horizontal_wrapped(|ui| {
                ui.add(egui::TextEdit::singleline(&mut self.smart.query)
                    .hint_text("Поиск по словам и формам...")
                    .desired_width(300.0));
                ui.label("Мин. фраз:");
                ui.add(egui::DragValue::new(&mut self.smart.min_count).range(1..=100000).speed(1));
                if ui.small_button("Выбрать найденные").clicked() {
                    for &i in &filtered { self.smart.selected.insert(self.smart.groups[i].key.clone()); }
                }
                if ui.small_button("Снять выделение").clicked() { self.smart.selected.clear(); }
            });
            ui.horizontal_wrapped(|ui| {
                if ui.add_enabled(selected_count > 0, egui::Button::new("+ В стоп-слова (скрыть)")).clicked() {
                    add_stop = true;
                }
                if ui.add_enabled(selected_count > 0, egui::Button::new("Удалить совпадающие фразы...")).clicked() {
                    delete = true;
                }
                if ui.button("Список стоп-слов").clicked() { show_stop_words = true; }
                ui.label(egui::RichText::new(
                    "Галочки слева — массовый выбор; клик по слову — просмотр справа."
                ).small().color(egui::Color32::from_gray(144)));
            });
            if !self.smart.feedback.is_empty() {
                ui.label(egui::RichText::new(&self.smart.feedback).color(egui::Color32::from_rgb(134, 204, 163)));
            }
            ui.separator();

            // Раздельные прокручиваемые панели: частые слова и ключевые фразы.
            let groups = &self.smart.groups;
            let selected = &mut self.smart.selected;
            let focused = &mut self.smart.focused;
            let original_rows = &self.rows;
            ui.columns(2, |columns| {
                columns[0].heading(format!("Группы слов ({})", filtered.len()));
                columns[0].label(egui::RichText::new("Отсортировано по числу совпадающих фраз ↓")
                    .small().color(egui::Color32::from_gray(143)));
                let max_height = columns[0].available_height().max(180.0);
                egui::ScrollArea::vertical()
                    .id_salt("smart_groups_scroll")
                    .max_height(max_height)
                    .show_rows(&mut columns[0], 27.0, filtered.len(), |ui, range| {
                        for pos in range {
                            let group = &groups[filtered[pos]];
                            let is_current = focused.as_ref() == Some(&group.key);
                            ui.horizontal(|ui| {
                                let mut checked = selected.contains(&group.key);
                                if ui.checkbox(&mut checked, "").changed() {
                                    if checked { selected.insert(group.key.clone()); }
                                    else { selected.remove(&group.key); }
                                }
                                let label = group.forms.iter().take(5).cloned().collect::<Vec<_>>().join(", ");
                                let label = if group.forms.len() > 5 { format!("{label}…") } else { label };
                                let response = ui.selectable_label(is_current, label);
                                if response.clicked() { *focused = Some(group.key.clone()); }
                                response.on_hover_text(group.forms.join(", "));
                                ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                                    ui.monospace(group.rows.len().to_string());
                                });
                            });
                        }
                    });

                columns[1].heading(if selected_count > 0 {
                    format!("Фразы из отмеченных групп ({})", preview_rows.len())
                } else {
                    format!("Фразы выбранного слова ({})", preview_rows.len())
                });
                if selected_count == 0 {
                    if let Some(group) = focused.as_ref()
                        .and_then(|key| groups.iter().find(|g| &g.key == key)) {
                        columns[1].label(egui::RichText::new(format!("Словоформы: {}", group.forms.join(", ")))
                            .small().color(egui::Color32::from_rgb(150, 196, 249)));
                    }
                } else {
                    columns[1].label(egui::RichText::new(
                        "Совпадения объединены без повторов — строка учитывается один раз."
                    ).small().color(egui::Color32::from_gray(143)));
                }
                let max_height = columns[1].available_height().max(180.0);
                egui::ScrollArea::vertical()
                    .id_salt("smart_phrases_scroll")
                    .max_height(max_height)
                    .show_rows(&mut columns[1], 27.0, preview_rows.len(), |ui, range| {
                        for pos in range {
                            let idx = preview_rows[pos];
                            if let Some(row) = original_rows.get(idx) {
                                ui.horizontal(|ui| {
                                    ui.label(egui::RichText::new(format!("{}", idx + 1))
                                        .small().color(egui::Color32::from_gray(120)));
                                    ui.label(row.phrase());
                                });
                            }
                        }
                    });
            });
        });

        if reanalyze { self.rebuild_smart_groups(); }
        if show_stop_words { self.show_stop_words = true; }
        if add_stop { self.smart_add_stop_words(); }
        if delete { self.smart_delete_phrases(); }
    }

    // ---------- Bukvarix ----------
    fn start_bukvarix(&mut self, ctx: &egui::Context) {
        let seeds: Vec<String> = self
            .buk
            .input
            .lines()
            .map(|l| l.trim().to_string())
            .filter(|l| !l.is_empty())
            .collect();
        if seeds.is_empty() {
            info("Информация", "Введите хотя бы одно ключевое слово.");
            return;
        }
        if seeds.len() > 100 {
            error("Можно ввести не более 100 ключевых слов за раз.");
            return;
        }
        let (tx, rx) = mpsc::channel();
        self.buk.rx = Some(rx);
        self.buk.busy = true;
        let key = self.buk.api_key.clone();
        let num = self.buk.num;
        let ctx = ctx.clone();
        std::thread::spawn(move || {
            let _ = tx.send(bukvarix_request(&seeds, &key, num));
            ctx.request_repaint();
        });
    }

    fn poll_bukvarix(&mut self) {
        let res = self.buk.rx.as_ref().map(|rx| rx.try_recv());
        match res {
            Some(Ok(r)) => {
                self.buk.rx = None;
                self.buk.busy = false;
                match r {
                    Ok(rows) => self.apply_bukvarix(rows),
                    Err(e) => error(&e),
                }
            }
            Some(Err(TryRecvError::Disconnected)) => {
                self.buk.rx = None;
                self.buk.busy = false;
            }
            _ => {}
        }
    }

    fn ensure_col(&mut self, name: &str) -> usize {
        if let Some(i) = self.columns.iter().position(|c| c == name) {
            return i;
        }
        self.columns.push(name.to_string());
        for r in &mut self.rows {
            r.cells.push(String::new());
        }
        self.columns.len() - 1
    }

    fn apply_bukvarix(&mut self, new: Vec<BukRow>) {
        if new.is_empty() {
            info("Результат", "Bukvarix не вернул данные по заданным ключам.");
            return;
        }
        if !self.has_data {
            self.columns = DEFAULT_COLS.iter().map(|s| s.to_string()).collect();
            self.rows.clear();
            self.has_data = true;
        }
        let norm = |s: &str| s.trim().to_lowercase().split_whitespace().collect::<String>();
        let find = |cols: &[String], needs: &[&str]| -> Option<usize> {
            cols.iter().position(|c| {
                let cn = norm(c);
                needs.iter().any(|n| cn == *n || cn.starts_with(n))
            })
        };
        let base = find(&self.columns, &["базоваячастотность"]);
        let exact = find(&self.columns, &["точнаячастотность"]);
        let very = find(&self.columns, &["оченьточнаячастотность", "оченьточнчастотность"]);
        let (ib, ie, iv) = match (base, exact, very) {
            (Some(b), Some(e), Some(v)) => (b, e, v),
            _ => (
                self.ensure_col("Частотность"),
                self.ensure_col("!Частостность"),
                self.ensure_col("[!Частостность]"),
            ),
        };
        let n = self.columns.len();
        let count = new.len();
        for b in new {
            let mut cells = vec![String::new(); n];
            cells[0] = b.phrase;
            cells[ib] = b.broad;
            cells[ie] = b.exact.clone();
            cells[iv] = b.exact; // в Bukvarix нет «очень точной» — дублируем exact
            self.rows.push(Row { group: String::new(), cells });
        }
        self.search.clear();
        self.group_filter = GroupFilter::All;
        self.refresh_all();
        self.buk.open = false;
        info("Успех", &format!("Добавлено строк из Bukvarix: {count}"));
    }

    // ===========================
    //  UI
    // ===========================
    fn top_panel(&mut self, ctx: &egui::Context) {
        egui::TopBottomPanel::top("workspace_header")
            .frame(egui::Frame::none()
                .fill(egui::Color32::from_rgb(22, 29, 39))
                .inner_margin(egui::Margin::symmetric(14.0, 10.0)))
            .show(ctx, |ui| {
                ui.horizontal(|ui| {
                    ui.label(egui::RichText::new("Clear Words")
                        .strong().size(17.0).color(egui::Color32::from_rgb(105, 176, 255)));
                    ui.add_space(12.0);
                    let project = self.project_path.as_ref()
                        .and_then(|p| p.file_name())
                        .map(|n| n.to_string_lossy().to_string())
                        .unwrap_or_else(|| "Новый проект".to_string());
                    ui.label(egui::RichText::new(project).color(egui::Color32::from_gray(173)));
                    ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                        if ui.button("Экспорт").clicked() { self.save_file(); }
                        if ui.button("Сохранить").clicked() { self.save_project(); }
                        ui.menu_button("Файл", |ui| {
                            if ui.button("Импорт CSV / Excel").clicked() { self.load_data(); ui.close_menu(); }
                            if ui.button("Открыть проект").clicked() { self.open_project(); ui.close_menu(); }
                            if ui.button("Сохранить проект как...").clicked() { self.save_project_as(); ui.close_menu(); }
                            if ui.button("Выгрузить XLSX / CSV").clicked() { self.save_file(); ui.close_menu(); }
                        });
                    });
                });
                ui.add_space(8.0);
                ui.horizontal_wrapped(|ui| {
                    for (tab, label) in [
                        (WorkspaceTab::Files, "Данные"),
                        (WorkspaceTab::Collect, "Сбор ключей"),
                        (WorkspaceTab::Clean, "Обработка"),
                        (WorkspaceTab::Groups, "Группировка"),
                        (WorkspaceTab::Smart, "Умная группировка"),
                        (WorkspaceTab::Export, "Экспорт"),
                    ] {
                        // Four spaces reserve room for a vector icon inside the same clickable tab.
                        let selected = self.tab == tab;
                        let response = ui.selectable_value(&mut self.tab, tab, format!("    {label}"));
                        let origin = egui::pos2(response.rect.left() + 7.0, response.rect.center().y - 7.0);
                        let color = if selected {
                            egui::Color32::from_rgb(178, 214, 255)
                        } else {
                            egui::Color32::from_gray(165)
                        };
                        paint_nav_icon(ui.painter(), origin, tab, color);
                    }
                    ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                        if ui.small_button("О программе").clicked() { self.show_contact = true; }
                    });
                });
                ui.separator();
                ui.horizontal_wrapped(|ui| {
                    match self.tab {
                        WorkspaceTab::Files => {
                            if ui.button("+ Импорт таблицы").clicked() { self.load_data(); }
                            if ui.button("Открыть проект").clicked() { self.open_project(); }
                            if ui.button("Сохранить как...").clicked() { self.save_project_as(); }
                            ui.label(egui::RichText::new("CSV, XLSX и собственные проекты .el")
                                .small().color(egui::Color32::from_gray(138)));
                        }
                        WorkspaceTab::Collect => {
                            if ui.button("+ Парсинг Bukvarix").clicked() { self.buk.open = true; }
                            if ui.button("Импорт CSV / Excel").clicked() { self.load_data(); }
                            if self.buk.busy { ui.spinner(); ui.label("Получаем ключевые слова..."); }
                        }
                        WorkspaceTab::Clean => {
                            if ui.button("Умная группировка").clicked() {
                                self.tab = WorkspaceTab::Smart;
                            }
                            ui.menu_button("Стоп-слова", |ui| {
                                if ui.button("Просмотреть / удалить").clicked() { self.show_stop_words = true; ui.close_menu(); }
                                if ui.button("Загрузить TXT").clicked() { self.load_stop_words_from_file(); ui.close_menu(); }
                                if ui.button("Сохранить TXT").clicked() { self.save_stop_words_to_file(); ui.close_menu(); }
                            });
                            ui.menu_button("Сортировка", |ui| {
                                if ui.button("По алфавиту").clicked() { self.sort_by(SortCol::Cell(0), true); ui.close_menu(); }
                                for (label, idx) in [("Частотность", 1), ("!Частостность", 2), ("[!Частостность]", 3)] {
                                    if ui.button(label).clicked() {
                                        if self.columns.len() > idx { self.sort_by(SortCol::Cell(idx), false); }
                                        ui.close_menu();
                                    }
                                }
                            });
                            ui.menu_button("Удалить дубли", |ui| {
                                if ui.button("Явные (по фразе)").clicked() {
                                    self.dedupe(norm_exact, "Удалены явные дубли",
                                        "сравнение по фразе без учета регистра/лишних пробелов/ё.");
                                    ui.close_menu();
                                }
                                if ui.button("Неявные (по составу слов)").clicked() {
                                    self.dedupe(norm_tokensort, "Удалены неявные дубли",
                                        "игнорируется пунктуация и порядок слов (слова сортируются).");
                                    ui.close_menu();
                                }
                            });
                            if ui.button("Назад").clicked() { self.undo_last_action(); }
                        }
                        WorkspaceTab::Groups => {
                            if ui.button("Анализ слов").clicked() {
                                self.tab = WorkspaceTab::Smart;
                            }
                            if ui.button("+ Новая ветка").clicked() { self.new_group = Some(String::new()); }
                            if ui.button("Перенести выделенные").clicked() { self.move_selected_to_group(); }
                            if ui.button("Перенести все видимые").clicked() { self.move_all_view_to_group(); }
                            if ui.button("Удалить ветку").clicked() { self.delete_group(); }
                        }
                        WorkspaceTab::Smart => {
                            ui.label("Группы слов — поиск стоп-слов и массовая очистка фраз");
                        }
                        WorkspaceTab::Export => {
                            if ui.button("Экспорт XLSX / CSV").clicked() { self.save_file(); }
                            if ui.button("Сохранить проект .el").clicked() { self.save_project(); }
                            if ui.button("Сохранить стоп-слова").clicked() { self.save_stop_words_to_file(); }
                        }
                    }
                });
            });
    }

    fn groups_panel(&mut self, ctx: &egui::Context) {
        let mut new_filter: Option<GroupFilter> = None;
        let mut dropped: Option<(GroupFilter, Vec<String>)> = None;
        let tree = build_group_tree(&self.groups, &self.rows);
        let ungrouped = self.rows.iter().filter(|r| r.group.trim().is_empty()).count();
        egui::SidePanel::left("group_navigation")
            .resizable(true)
            .default_width(238.0)
            .width_range(185.0..=360.0)
            .frame(egui::Frame::none()
                .fill(egui::Color32::from_rgb(25, 32, 43))
                .inner_margin(egui::Margin::same(12.0)))
            .show(ctx, |ui| {
                ui.horizontal(|ui| {
                    ui.heading("Группы");
                    ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                        if ui.small_button("+ Новая").clicked() { self.new_group = Some(String::new()); }
                    });
                });
                ui.add_space(4.0);
                ui.add(egui::TextEdit::singleline(&mut self.group_search)
                    .hint_text("Найти группу...")
                    .desired_width(f32::INFINITY));
                ui.add_space(8.0);
                let available_tree_height = (ui.available_height() - 110.0).max(120.0);
                egui::ScrollArea::vertical()
                    .id_salt("group_tree_scroll")
                    .max_height(available_tree_height)
                    .show(ui, |ui| {
                    let all = GroupFilter::All;
                    let response = ui.selectable_label(self.group_filter == all,
                        format!("Все ключи   {}", self.rows.len()));
                    if response.clicked() { new_filter = Some(all.clone()); }
                    if let Some(payload) = response.dnd_release_payload::<DragPayload>() {
                        dropped = Some((all, payload.0.clone()));
                    }
                    let no = GroupFilter::NoGroup;
                    let response = ui.selectable_label(self.group_filter == no,
                        format!("Без группы   {ungrouped}"));
                    if response.clicked() { new_filter = Some(no.clone()); }
                    if let Some(payload) = response.dnd_release_payload::<DragPayload>() {
                        dropped = Some((no, payload.0.clone()));
                    }
                    ui.add_space(10.0);
                    ui.label(egui::RichText::new("МОИ ВЕТКИ")
                        .small().strong().color(egui::Color32::from_gray(147)));
                    ui.add_space(4.0);
                    let search = self.group_search.trim().to_lowercase();
                    if tree.is_empty() {
                        ui.label(egui::RichText::new("Пока нет групп. Нажми «+ Новая».")
                            .color(egui::Color32::from_gray(145)));
                    }
                    for node in tree.values() {
                        draw_group_node(ui, node, &self.group_filter, &search, &mut new_filter, &mut dropped);
                    }
                });
                ui.separator();
                ui.horizontal_wrapped(|ui| {
                    if ui.button("Перенести выбранное").clicked() { self.move_selected_to_group(); }
                    if ui.small_button("Удалить ветку").clicked() { self.delete_group(); }
                });
                ui.add_space(5.0);
                ui.label(egui::RichText::new(
                    "Ctrl / Shift — выделить строки. Перетащи их на ветку слева."
                ).small().color(egui::Color32::from_gray(146)));
            });
        if let Some(f) = new_filter {
            self.group_filter = f;
            self.recompute_view();
        }
        if let Some((target, phrases)) = dropped {
            if self.has_data && !phrases.is_empty() {
                match target {
                    GroupFilter::All => self.open_picker(phrases),
                    GroupFilter::NoGroup => self.assign_group(&phrases, ""),
                    GroupFilter::Named(g) => self.assign_group(&phrases, &g),
                }
            }
        }
    }

    fn details_panel(&mut self, ctx: &egui::Context) {
        if !self.show_details { return; }
        let first = self.selected.iter().next()
            .and_then(|p| self.rows.iter().find(|r| r.phrase() == p))
            .cloned();
        egui::SidePanel::right("selection_details")
            .default_width(246.0).resizable(true)
            .width_range(200.0..=340.0)
            .show(ctx, |ui| {
                ui.horizontal(|ui| {
                    ui.heading("Детали");
                    ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                        if ui.small_button("✕").clicked() { self.show_details = false; }
                    });
                });
                ui.separator();
                ui.label(format!("Выбрано строк: {}", self.selected.len()));
                if let Some(row) = first {
                    ui.add_space(9.0);
                    ui.label(egui::RichText::new(row.phrase()).strong().size(16.0));
                    ui.label(egui::RichText::new(if row.group.is_empty() {
                        "Без группы"
                    } else { row.group.as_str() }).color(egui::Color32::from_gray(153)));
                    ui.add_space(7.0);
                    for (i, value) in row.cells.iter().enumerate().skip(1) {
                        if let Some(label) = self.columns.get(i) {
                            ui.horizontal_wrapped(|ui| {
                                ui.label(egui::RichText::new(label).color(egui::Color32::from_gray(159)));
                                ui.label(egui::RichText::new(value).strong());
                            });
                        }
                    }
                    ui.add_space(11.0);
                    if ui.button("Перенести в группу...").clicked() { self.move_selected_to_group(); }
                    if ui.button("Удалить выбранные").clicked() { self.delete_selected_rows(); }
                    ui.add_space(9.0);
                    ui.label(egui::RichText::new(
                        "Двойной щелчок по строке — выбор стоп-слова."
                    ).small().color(egui::Color32::from_gray(140)));
                } else {
                    ui.add_space(15.0);
                    ui.label(egui::RichText::new("Выбери строку таблицы, чтобы увидеть информацию.")
                        .color(egui::Color32::from_gray(152)));
                }
            });
    }

    fn table_panel(&mut self, ctx: &egui::Context) {
        let mut sort_click: Option<usize> = None;
        let mut click: Option<(usize, egui::Modifiers)> = None;
        let mut dbl: Option<String> = None;
        let mut want_delete = false;
        let mut want_move = false;
        let mut rc_select: Option<String> = None;
        let mut select_only: Option<String> = None;

        egui::CentralPanel::default()
            .frame(egui::Frame::none().fill(egui::Color32::from_rgb(19, 25, 34))
                .inner_margin(egui::Margin::symmetric(14.0, 12.0)))
            .show(ctx, |ui| {
                ui.horizontal(|ui| {
                    ui.heading("Ключевые фразы");
                    ui.label(egui::RichText::new(format!("{} всего", self.rows.len()))
                        .color(egui::Color32::from_gray(146)));
                });
                ui.add_space(8.0);
                let grouped = self.rows.iter().filter(|r| !r.group.trim().is_empty()).count();
                ui.horizontal_wrapped(|ui| {
                    for (label, n) in [
                        ("Всего", self.rows.len()),
                        ("Показано", self.view.len()),
                        ("В группах", grouped),
                        ("Выбрано", self.selected.len()),
                    ] {
                        egui::Frame::none()
                            .fill(egui::Color32::from_rgb(31, 40, 53))
                            .rounding(egui::Rounding::same(6.0))
                            .inner_margin(egui::Margin::symmetric(12.0, 7.0))
                            .show(ui, |ui| {
                                ui.horizontal(|ui| {
                                    ui.label(egui::RichText::new(label)
                                        .color(egui::Color32::from_gray(163)));
                                    ui.label(egui::RichText::new(n.to_string())
                                        .color(egui::Color32::WHITE).strong());
                                });
                            });
                    }
                });
                ui.add_space(11.0);
                ui.horizontal_wrapped(|ui| {
                    ui.label("⌕");
                    let width = (ui.available_width() - 260.0).max(160.0);
                    if ui.add(egui::TextEdit::singleline(&mut self.search)
                        .hint_text("Поиск: все введённые слова должны присутствовать в фразе")
                        .desired_width(width)).changed() {
                        self.recompute_view();
                    }
                    if ui.button("Сбросить").clicked() {
                        self.search.clear();
                        self.recompute_view();
                    }
                    ui.menu_button("Колонки ▾", |ui| {
                        for (i, name) in self.columns.iter().enumerate() {
                            if i == 0 { continue; }
                            if i >= self.visible_columns.len() {
                                self.visible_columns.resize(self.columns.len(), true);
                            }
                            ui.checkbox(&mut self.visible_columns[i], name);
                        }
                    });
                    if ui.selectable_label(self.show_details, "Детали").clicked() {
                        self.show_details = !self.show_details;
                    }
                });
                ui.add_space(9.0);
                ui.separator();
                if self.view.is_empty() {
                    ui.add_space(65.0);
                    ui.vertical_centered(|ui| {
                        ui.label(egui::RichText::new(if self.has_data {
                            "Ничего не найдено"
                        } else { "Начни с импорта или сбора ключевых фраз" })
                            .size(19.0).strong());
                        ui.add_space(6.0);
                        ui.label(egui::RichText::new(if self.has_data {
                            "Проверь поиск, выбранную группу и стоп-слова"
                        } else { "Открой CSV/Excel или запусти парсинг Bukvarix" })
                            .color(egui::Color32::from_gray(151)));
                        ui.add_space(14.0);
                        ui.horizontal(|ui| {
                            ui.add_space((ui.available_width() - 275.0).max(0.0) / 2.0);
                            if ui.button("Импорт файла").clicked() { self.load_data(); }
                            if ui.button("Парсинг Bukvarix").clicked() { self.buk.open = true; }
                        });
                    });
                    return;
                }

                let rows = &self.rows;
                let view = &self.view;
                let selected = &self.selected;
                let columns = &self.columns;
                let visible: Vec<usize> = (0..columns.len())
                    .filter(|&i| i == 0 || self.visible_columns.get(i).copied().unwrap_or(true))
                    .collect();
                egui::ScrollArea::horizontal().auto_shrink([false, false]).show(ui, |ui| {
                    let mut tb = TableBuilder::new(ui)
                        .striped(true)
                        .resizable(true)
                        .sense(egui::Sense::click_and_drag())
                        .cell_layout(egui::Layout::left_to_right(egui::Align::Center))
                        .column(Column::initial(46.0).at_least(35.0).clip(true))
                        .column(Column::initial(405.0).at_least(180.0).clip(true))
                        .column(Column::initial(155.0).at_least(95.0).clip(true));
                    for _ in visible.iter().filter(|&&i| i != 0) {
                        tb = tb.column(Column::initial(150.0).at_least(95.0).clip(true));
                    }
                    tb.header(30.0, |mut header| {
                        header.col(|ui| { ui.label("№"); });
                        header.col(|ui| {
                            if ui.add(egui::Button::new(egui::RichText::new("Фраза ↕").strong())
                                .frame(false)).clicked() { sort_click = Some(1); }
                        });
                        header.col(|ui| {
                            if ui.add(egui::Button::new(egui::RichText::new("Группа ↕").strong())
                                .frame(false)).clicked() { sort_click = Some(0); }
                        });
                        for &i in visible.iter().filter(|&&i| i != 0) {
                            header.col(|ui| {
                                let name = &columns[i];
                                if ui.add(egui::Button::new(egui::RichText::new(format!("{name} ↕")).strong())
                                    .frame(false)).clicked() { sort_click = Some(i + 1); }
                            });
                        }
                    }).body(|body| {
                        body.rows(29.0, view.len(), |mut row| {
                            let pos = row.index();
                            let r = &rows[view[pos]];
                            let phrase = r.phrase();
                            let is_selected = selected.contains(phrase);
                            row.set_selected(is_selected);
                            row.col(|ui| {
                                ui.label(egui::RichText::new(if is_selected {
                                    "✓".to_string()
                                } else { (pos + 1).to_string() })
                                    .color(if is_selected {
                                        egui::Color32::from_rgb(108, 185, 255)
                                    } else { egui::Color32::from_gray(135) }));
                            });
                            row.col(|ui| {
                                ui.add(egui::Label::new(egui::RichText::new(phrase).color(
                                    egui::Color32::from_rgb(220, 228, 240)))
                                    .selectable(false));
                            });
                            row.col(|ui| {
                                ui.label(egui::RichText::new(if r.group.is_empty() { "—" } else { &r.group })
                                    .color(egui::Color32::from_gray(156)));
                            });
                            for &i in visible.iter().filter(|&&i| i != 0) {
                                row.col(|ui| {
                                    let value = r.cells.get(i).map(String::as_str).unwrap_or("");
                                    ui.add(egui::Label::new(value).selectable(false));
                                });
                            }
                            let response = row.response();
                            if response.double_clicked() {
                                dbl = Some(phrase.to_string());
                            } else if response.clicked() {
                                click = Some((pos, ctx.input(|i| i.modifiers)));
                            }
                            if response.drag_started() {
                                let phrases: Vec<String> = if selected.contains(phrase) {
                                    selected.iter().cloned().collect()
                                } else {
                                    select_only = Some(phrase.to_string());
                                    vec![phrase.to_string()]
                                };
                                egui::DragAndDrop::set_payload(ctx, DragPayload(phrases));
                            }
                            if response.dragged() { ctx.set_cursor_icon(egui::CursorIcon::Grabbing); }
                            if response.secondary_clicked() { rc_select = Some(phrase.to_string()); }
                            response.context_menu(|ui| {
                                if ui.button("Перенести в группу...").clicked() {
                                    want_move = true;
                                    ui.close_menu();
                                }
                                if ui.button("Удалить выбранное").clicked() {
                                    want_delete = true;
                                    ui.close_menu();
                                }
                            });
                        });
                    });
                });
            });

        if let Some(p) = select_only {
            self.selected.clear();
            self.selected.insert(p);
        }
        if let Some(p) = rc_select {
            if !self.selected.contains(&p) {
                self.selected.clear();
                self.selected.insert(p);
            }
        }
        if let Some((pos, m)) = click {
            let ph = self.rows[self.view[pos]].phrase().to_string();
            if m.command || m.ctrl {
                if !self.selected.remove(&ph) { self.selected.insert(ph); }
            } else if let (true, Some(a)) = (m.shift, self.last_clicked) {
                for p in a.min(pos)..=a.max(pos) {
                    self.selected.insert(self.rows[self.view[p]].phrase().to_string());
                }
            } else {
                self.selected.clear();
                self.selected.insert(ph);
            }
            self.last_clicked = Some(pos);
        }
        if let Some(c) = sort_click { self.sort_by_tree_column(c); }
        if let Some(p) = dbl { self.open_word_selection(&p); }
        if want_delete { self.delete_selected_rows(); }
        if want_move { self.move_selected_to_group(); }
    }

    fn windows(&mut self, ctx: &egui::Context) {
        // --- выбор слова для стоп-слов ---
        if let Some(words) = self.word_dialog.take() {
            let mut keep = true;
            let mut chosen: Option<String> = None;
            egui::Window::new("Выбор слова").open(&mut keep).collapsible(false).show(ctx, |ui| {
                egui::ScrollArea::vertical().max_height(300.0).show(ui, |ui| {
                    for w in &words {
                        if ui
                            .add_sized(
                                [ui.available_width(), 22.0],
                                egui::Button::new(format!("Добавить '{w}' в стоп-слова")),
                            )
                            .clicked()
                        {
                            chosen = Some(w.clone());
                        }
                    }
                });
            });
            if let Some(w) = chosen {
                self.add_stop_word(w);
            } else if keep {
                self.word_dialog = Some(words);
            }
        }

        // --- список стоп-слов ---
        if self.show_stop_words {
            let mut open = true;
            let mut remove: Option<String> = None;
            let sw = self.sorted_stop_words();
            egui::Window::new("Список стоп-слов").open(&mut open).show(ctx, |ui| {
                egui::ScrollArea::vertical().max_height(400.0).show(ui, |ui| {
                    for w in &sw {
                        ui.horizontal(|ui| {
                            ui.label(w);
                            ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                                if ui.button("Удалить").clicked() {
                                    remove = Some(w.clone());
                                }
                            });
                        });
                    }
                });
            });
            self.show_stop_words = open;
            if let Some(w) = remove {
                self.remove_stop_word(&w);
            }
        }

        // --- новая ветка ---
        if let Some(mut name) = self.new_group.take() {
            let mut ok = false;
            let mut cancel = false;
            egui::Window::new("Новая ветка").collapsible(false).show(ctx, |ui| {
                ui.label("Название ветки (например, SEO/Услуги для вложенной папки):");
                let r = ui.text_edit_singleline(&mut name);
                if r.lost_focus() && ui.input(|i| i.key_pressed(egui::Key::Enter)) {
                    ok = true;
                }
                ui.horizontal(|ui| {
                    if ui.button("OK").clicked() {
                        ok = true;
                    }
                    if ui.button("Отмена").clicked() {
                        cancel = true;
                    }
                });
            });
            if ok {
                let n = name.trim().to_string();
                if !n.is_empty() {
                    if self.groups.contains(&n) {
                        info("Информация", &format!("Ветка '{n}' уже существует."));
                    } else {
                        self.groups.push(n);
                        self.rebuild_groups();
                    }
                }
            } else if !cancel {
                self.new_group = Some(name);
            }
        }

        // --- выбор целевой ветки ---
        if let Some(mut p) = self.picker.take() {
            let mut ok = false;
            let mut cancel = false;
            let groups = self.groups.clone();
            egui::Window::new("Перенос в ветку").collapsible(false).show(ctx, |ui| {
                ui.label("Выберите ветку, в которую перенести выбранные фразы:");
                egui::ScrollArea::vertical().max_height(260.0).show(ui, |ui| {
                    if ui.selectable_label(p.sel == 0, "Без ветки").clicked() {
                        p.sel = 0;
                    }
                    for (i, g) in groups.iter().enumerate() {
                        if ui.selectable_label(p.sel == i + 1, g).clicked() {
                            p.sel = i + 1;
                        }
                    }
                });
                ui.horizontal(|ui| {
                    if ui.button("OK").clicked() {
                        ok = true;
                    }
                    if ui.button("Отмена").clicked() {
                        cancel = true;
                    }
                });
            });
            if ok {
                let target = if p.sel == 0 { String::new() } else { groups[p.sel - 1].clone() };
                self.last_target_group = Some(target.clone());
                self.assign_group(&p.phrases, &target);
            } else if !cancel {
                self.picker = Some(p);
            }
        }

        // --- контакты ---
        if self.show_contact {
            let mut open = true;
            egui::Window::new("Связаться с автором").open(&mut open).show(ctx, |ui| {
                ui.label("Обработка слов v1.2\nАвтор: Эльдар Ибрагимов");
                ui.horizontal(|ui| {
                    if ui.button("ВК").clicked() {
                        let _ = open::that("https://vk.com/mr.crutch");
                    }
                    if ui.button("TG").clicked() {
                        let _ = open::that("https://t.me/God_SMM");
                    }
                });
            });
            self.show_contact = open;
        }

        // --- Bukvarix: only dialog presentation was redesigned ---
        if self.buk.open {
            let mut open = true;
            let mut start = false;
            egui::Window::new("Сбор ключевых слов · Bukvarix")
                .open(&mut open)
                .collapsible(false)
                .resizable(true)
                .default_width(610.0)
                .show(ctx, |ui| {
                    ui.label(egui::RichText::new("Ключевые слова для поиска").strong().size(16.0));
                    ui.label(egui::RichText::new(
                        "До 100 фраз, каждая с новой строки. Результат добавится в текущую таблицу."
                    ).color(egui::Color32::from_gray(165)));
                    ui.add_space(8.0);
                    ui.add(egui::TextEdit::multiline(&mut self.buk.input)
                        .hint_text("seo продвижение\nseo оптимизация")
                        .desired_rows(10)
                        .desired_width(f32::INFINITY));
                    ui.add_space(8.0);
                    ui.horizontal_wrapped(|ui| {
                        ui.label("API-ключ");
                        ui.add(egui::TextEdit::singleline(&mut self.buk.api_key)
                            .desired_width(180.0));
                        ui.add_space(14.0);
                        ui.label("Строк в отчёте");
                        ui.add(egui::DragValue::new(&mut self.buk.num).range(10..=1_000_000));
                    });
                    ui.separator();
                    ui.horizontal(|ui| {
                        if ui.add_enabled(!self.buk.busy,
                            egui::Button::new("Начать парсинг")).clicked() {
                            start = true;
                        }
                        if self.buk.busy {
                            ui.spinner();
                            ui.label("Ожидаем ответ Bukvarix...");
                        }
                    });
                });
            self.buk.open = open;
            if start { self.start_bukvarix(ctx); }
        }

    }
}

fn apply_ui_theme(ctx: &egui::Context) {
    let mut visuals = egui::Visuals::dark();
    visuals.panel_fill = egui::Color32::from_rgb(19, 25, 34);
    visuals.window_fill = egui::Color32::from_rgb(29, 37, 49);
    visuals.extreme_bg_color = egui::Color32::from_rgb(15, 20, 28);
    visuals.widgets.inactive.bg_fill = egui::Color32::from_rgb(40, 51, 66);
    visuals.widgets.hovered.bg_fill = egui::Color32::from_rgb(52, 69, 91);
    visuals.selection.bg_fill = egui::Color32::from_rgb(37, 91, 151);
    ctx.set_visuals(visuals);
    let mut style = (*ctx.style()).clone();
    style.spacing.item_spacing = egui::vec2(8.0, 7.0);
    style.spacing.button_padding = egui::vec2(10.0, 6.0);
    ctx.set_style(style);
}

impl eframe::App for App {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        self.poll_bukvarix();
        self.top_panel(ctx);
        egui::TopBottomPanel::bottom("workspace_status")
            .frame(egui::Frame::none()
                .fill(egui::Color32::from_rgb(22, 29, 39))
                .inner_margin(egui::Margin::symmetric(12.0, 6.0)))
            .show(ctx, |ui| {
                ui.horizontal(|ui| {
                    ui.label(format!("Всего: {}     Показано: {}     Выделено: {}",
                        self.rows.len(), self.view.len(), self.selected.len()));
                    ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                        ui.label(if self.buk.busy { "Bukvarix: выполняется запрос" }
                            else if self.tab == WorkspaceTab::Smart { "Анализ слов · Стоп-слова · Массовая очистка" }
                            else { "Ctrl / Shift · Перетаскивание · Двойной клик" });
                    });
                });
            });
        if self.tab == WorkspaceTab::Smart {
            self.smart_group_panel(ctx);
        } else {
            self.groups_panel(ctx);
            self.details_panel(ctx);
            self.table_panel(ctx);
        }
        self.windows(ctx);
    }
}

fn main() -> eframe::Result<()> {
    let options = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default()
            .with_inner_size([1400.0, 850.0])
            .with_min_inner_size([900.0, 590.0])
            .with_title("Clear Words — рабочая область"),
        ..Default::default()
    };
    eframe::run_native(
        "Clear Words",
        options,
        Box::new(|cc| {
            apply_ui_theme(&cc.egui_ctx);
            Ok(Box::new(App::new()))
        }),
    )
}


#[cfg(test)]
mod smart_tests {
    use super::*;

    fn phrase(text: &str) -> Row {
        Row { group: String::new(), cells: vec![text.to_owned()] }
    }

    #[test]
    fn exact_grouping_counts_distinct_rows() {
        let rows = vec![phrase("SEO seo seo"), phrase("Seo курсы"), phrase("другой запрос")];
        let groups = build_smart_groups(&rows, &[0, 1, 2], false);
        let seo = groups.iter().find(|g| g.forms == vec!["seo"]).unwrap();
        assert_eq!(seo.rows.len(), 2); // три вхождения в строке != три фразы
    }

    #[test]
    fn stems_group_russian_word_forms() {
        let rows = vec![phrase("продвижение сайта"), phrase("цена продвижения")];
        let groups = build_smart_groups(&rows, &[0, 1], true);
        let group = groups.iter().find(|g| g.forms.contains(&"продвижение".to_owned())).unwrap();
        assert!(group.forms.contains(&"продвижения".to_owned()));
        assert_eq!(group.rows.len(), 2);
    }

    #[test]
    fn no_substring_matches_and_scope_respected() {
        let rows = vec![phrase("seo"), phrase("seotools"), phrase("seo агентство")];
        let groups = build_smart_groups(&rows, &[0, 1], false);
        let seo = groups.iter().find(|g| g.forms == vec!["seo"]).unwrap();
        assert_eq!(seo.rows, vec![0]);
    }
}
