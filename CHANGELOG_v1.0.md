# Tool Spy Idea v1.0.0 — Session Improvements

Tổng cộng **8 features mới** thêm vào bản v1.0.0.

---

## 🎨 1. Dark/Light Theme Toggle

- Switch theme trong sidebar (button "Sáng/Tối") hoặc Settings → Giao diện
- Lưu vào `localStorage` + sync server
- Áp dụng `data-theme="light"` trên `<html>` với 40+ CSS overrides
- Tự load theme đã chọn lần sau

## 🌐 2. Vietnamese / English Toggle

- Switch ngôn ngữ trong sidebar (button "VI/EN") hoặc Settings → Ngôn ngữ
- I18N dictionary 50+ keys trong `I18N.vi` / `I18N.en`
- Auto-translate: sidebar, header, tabs, modals, buttons, descriptions
- Tab content nội bộ vẫn tiếng Việt (có thể mở rộng sau)

## 📊 3. Stats Dashboard Sidebar

- 4 counter real-time: Spied / Downloaded / Cleaned / Dropbox Links
- Tự động tăng khi job complete
- Lưu vào `data/config.json` → persist qua sessions
- Reset được từ Settings → Stats & Updates

## 🕐 4. Spy History

- Nút "Lịch sử spy" ở sidebar (có badge đếm)
- Modal overlay với danh sách history mới nhất trước
- Mỗi entry: icon màu theo loại + label + timestamp + count
- Delete từng entry hoặc Clear all
- Lưu `data/spy_history.json` (max 100 entries gần nhất)
- Hotkey **Ctrl+H** toggle mở/đóng, **ESC** đóng

## ↩️ 5. Undo/Redo Clean Title

- 2 button Undo/Redo bên cạnh Export Excel
- Hotkey **Ctrl+Z** (undo), **Ctrl+Y** / **Ctrl+Shift+Z** (redo)
- Stack 30 snapshot gần nhất, reset mỗi lần process mới
- Track mọi edit cell fixed của title

## 🔄 6. Auto-Update Check

- Silent check lúc mở app, show badge 🟢 trên Settings nếu có bản mới
- Manual check trong Settings → Version
- URL check configurable ở `UPDATE_CHECK_URL` trong `main.py`
- Để trống = không check (default)
- Khi có update: hiển thị version mới + link download + notes

## 📁 7. Export CSV

- Thêm nút **CSV** bên cạnh Export Excel trong: Spy Link, Etsy Shop, Clean Title
- BOM UTF-8 nên Excel + Google Sheets nhận đúng tiếng Việt
- Format giống Excel: cùng column layout

## 🎯 8. Drag & Drop File

- Kéo thả file vào textarea URL (Spy tab) → import URLs
- Kéo thả file vào textarea Titles (Clean tab) → import titles
- Kéo thả file vào nút Import Excel (Download tab) → import products
- Hỗ trợ: `.txt`, `.csv`, `.xlsx`, `.xls`
- Visual feedback: "📂 Thả file vào đây" overlay khi hover

---

## 🔧 Technical Changes

### Backend (`main.py`)
- `APP_VERSION = "1.0.0"`, `HISTORY_FILE`, `UPDATE_CHECK_URL` constants
- `load_config()` mở rộng với `theme/lang/stats` defaults
- Helper functions: `load_history()`, `save_history()`, `add_history_entry()`, `bump_stat()`
- **10 endpoints mới**:
  - `GET /api/history/list`, `POST /api/history/add`, `POST /api/history/delete/<id>`, `POST /api/history/clear`
  - `GET /api/stats`, `POST /api/stats/bump`, `POST /api/stats/reset`
  - `GET /api/version` — check update
  - `POST /api/spy/export-csv`, `/api/spy/etsy-shop-export-csv`, `/api/clean/export-csv`, `/api/dropbox/export-csv`

### Frontend CSS (`static/index.html`)
- Light theme overrides: 40+ rules dưới `:root[data-theme="light"]`
- Components: `.stat-pill`, `.history-item`, `.tool-toggle-btn`, `.update-badge`
- Existing `.drop-zone.drag-over` CSS đã sẵn, giờ được activate bởi JS handlers

### Frontend JS
- `I18N` dictionary với 50+ keys × 2 ngôn ngữ
- Global refs: `_globalT`, `_globalLang`, `_globalBumpStat`, `_globalAddHistory`
- Helpers: `bumpStat()`, `pushHistory()`, `downloadBlob()`, `useDrop()` hook
- App state mới: `theme`, `lang`, `stats`, `history`, `historyOpen`, `updateInfo`, `checkingUpdate`
- Clean state thêm: `undoStack`, `redoStack` (max 30)

### Fixes dọc đường
- Fix bug Dropbox polling có `if-else` sai cấu trúc (else sau ternary never executed)

---

## 📦 Files mới/sửa

Sửa: `main.py`, `static/index.html`

Mới tạo khi chạy (data dir):
- `data/spy_history.json` — lưu history
- `data/config.json` — lưu stats + theme/lang

---

## 🎮 Hotkey Reference

| Hotkey | Action |
|--------|--------|
| `Ctrl+1` → `Ctrl+5` | Switch tab |
| `Ctrl+H` | Toggle History modal |
| `Ctrl+Z` | Undo (trong Clean tab) |
| `Ctrl+Y` / `Ctrl+Shift+Z` | Redo (trong Clean tab) |
| `ESC` | Close modal (lightbox / history) |

---

## 🛠️ Cấu hình Auto-Update

Nếu bạn muốn enable update check, sửa trong `main.py`:

```python
UPDATE_CHECK_URL = "https://raw.githubusercontent.com/YOUR_USERNAME/toolspyidea/main/version.json"
```

File `version.json` trên GitHub có format:
```json
{
  "version": "1.1.0",
  "notes": "Added feature X",
  "download_url": "https://github.com/.../releases/download/v1.1.0/Setup.exe"
}
```

---

**Session code by Claude · v1.0.0 polish pass**
