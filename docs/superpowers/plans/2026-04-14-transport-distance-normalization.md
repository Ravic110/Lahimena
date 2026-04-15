# Transport Distance Normalization Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Corriger l'affichage/calcul de distance dans la cotation transport client, fiabiliser les recherches `KM_MADA`, et nettoyer les données existantes de `INFOS_CLIENTS`.

**Architecture:** La correction repose sur une normalisation centralisée des repères dans `utils/excel_handler.py`, un lookup `KM_MADA` robuste face aux doublons, puis l'alignement de `gui/forms/client_transport_cotation.py` sur le calcul existant du module transport principal. Le nettoyage des données existantes est encapsulé dans une fonction dédiée pour rester testable et réutilisable.

**Tech Stack:** Python, pytest, openpyxl, CustomTkinter/Tkinter, Excel workbooks `data.xlsx` et `data-hotel.xlsx`

---

### Task 1: Ajouter les tests rouges sur la normalisation et le lookup `KM_MADA`

**Files:**
- Modify: `tests/test_excel_handler.py`
- Modify: `utils/excel_handler.py`
- Test: `tests/test_excel_handler.py`

- [ ] **Step 1: Write the failing tests**

```python
def test_normalize_km_mada_repere_removes_duration_suffixes():
    assert normalize_km_mada_repere("Antananarivo(1 jours)") == "antananarivo"
    assert normalize_km_mada_repere(" Antsirabe (2 jours) ") == "antsirabe"


def test_normalize_km_mada_repere_applies_business_aliases():
    assert normalize_km_mada_repere("Tuler") == "toliary"
    assert normalize_km_mada_repere("Tulear") == "toliary"
    assert normalize_km_mada_repere("Ranohira (Isalo)") == "ranohira"


def test_get_km_mada_km_for_repere_prefers_positive_duplicate(monkeypatch):
    rows = [
        {"repere": "MORONDAVA", "km": 679, "duree": 5},
        {"repere": "MORONDAVA", "km": 0, "duree": 0},
    ]
    monkeypatch.setattr("utils.excel_handler._load_km_mada_rows", lambda: rows)
    _invalidate_km_mada_cache()

    assert get_km_mada_km_for_repere("Morondava") == 679
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `./.env/bin/pytest tests/test_excel_handler.py -k "normalize_km_mada_repere or prefers_positive_duplicate" -v`

Expected: FAIL because `normalize_km_mada_repere` does not exist yet and duplicate selection is not implemented.

- [ ] **Step 3: Implement the minimal helpers and duplicate-selection logic**

```python
_KM_MADA_ALIASES = {
    "tuler": "toliary",
    "tulear": "toliary",
    "ranohira isalo": "ranohira",
}


def normalize_km_mada_repere(value):
    text = str(value or "").strip().lower()
    text = re.sub(r"\(\s*\d+\s*jour[s]?\s*\)", " ", text)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"[^a-z0-9]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return _KM_MADA_ALIASES.get(text, text)


def _select_best_km_mada_row(rows):
    positive = [row for row in rows if _parse_num(row.get("km", 0)) > 0]
    if positive:
        return max(positive, key=lambda row: _parse_num(row.get("km", 0)))
    return rows[0] if rows else None
```

- [ ] **Step 4: Update the cache and lookup to use normalized keys**

```python
lookup = {}
for row in rows:
    key = normalize_km_mada_repere(row.get("repere"))
    if key:
        lookup.setdefault(key, []).append(row)


def get_km_mada_km_for_repere(repere):
    key = normalize_km_mada_repere(repere)
    candidates = _KM_MADA_CACHE["lookup"].get(key, [])
    selected = _select_best_km_mada_row(candidates)
    return _parse_num(selected.get("km", 0)) if selected else 0
```

- [ ] **Step 5: Run the tests to verify they pass**

Run: `./.env/bin/pytest tests/test_excel_handler.py -k "normalize_km_mada_repere or prefers_positive_duplicate" -v`

Expected: PASS

- [ ] **Step 6: Commit**

```bash
git add tests/test_excel_handler.py utils/excel_handler.py
git commit -m "Fix KM_MADA normalization and duplicate lookup"
```

### Task 2: Ajouter les tests rouges sur la distance de segment côté cotation transport client

**Files:**
- Create: `tests/test_client_transport_cotation.py`
- Modify: `gui/forms/client_transport_cotation.py`
- Test: `tests/test_client_transport_cotation.py`

- [ ] **Step 1: Write the failing tests**

```python
from gui.forms.client_transport_cotation import _make_segments, _compute_segment_km


def test_make_segments_normalizes_existing_client_route_values():
    client = {
        "ville_depart": "Antananarivo(1 jours)",
        "ville_arrivee": "Antsirabe(2 jours), Ranomafana(2 jours)",
        "itineraire_circuit": "",
    }

    assert _make_segments(client) == [
        ("Antananarivo", "Antsirabe"),
        ("Antsirabe", "Ranomafana"),
    ]


def test_compute_segment_km_uses_departure_and_arrival_difference(monkeypatch):
    km_values = {
        "antananarivo": 0,
        "antsirabe": 169,
        "ranomafana": 235,
    }

    monkeypatch.setattr(
        "gui.forms.client_transport_cotation.get_km_mada_km_for_repere",
        lambda city: km_values.get(city.lower(), 0),
    )

    assert _compute_segment_km("Antananarivo", "Antsirabe") == 169
    assert _compute_segment_km("Antsirabe", "Ranomafana") == 66
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `./.env/bin/pytest tests/test_client_transport_cotation.py -v`

Expected: FAIL because `_compute_segment_km` does not exist yet and `_make_segments()` keeps raw city values.

- [ ] **Step 3: Implement minimal segment normalization helpers**

```python
from utils.excel_handler import normalize_km_mada_repere


def _display_city_name(value: str) -> str:
    normalized = normalize_km_mada_repere(value)
    return " ".join(part.capitalize() for part in normalized.split())


def _compute_segment_km(depart: str, arrivee: str) -> int:
    km_dep = float(get_km_mada_km_for_repere(depart) or 0)
    km_arr = float(get_km_mada_km_for_repere(arrivee) or 0)
    distance = round(abs(km_arr - km_dep))
    return int(distance) if distance > 0 else 0
```

- [ ] **Step 4: Update `_parse_cities()` and `_make_segments()` to emit cleaned display values**

```python
clean_city = _display_city_name(city)
if clean_city and _normalize(clean_city) not in seen:
    seen.add(_normalize(clean_city))
    result.append(clean_city)
```

- [ ] **Step 5: Run the tests to verify they pass**

Run: `./.env/bin/pytest tests/test_client_transport_cotation.py -v`

Expected: PASS

- [ ] **Step 6: Commit**

```bash
git add tests/test_client_transport_cotation.py gui/forms/client_transport_cotation.py
git commit -m "Fix client transport segment distance calculation"
```

### Task 3: Brancher le nouveau calcul dans l'écran client transport

**Files:**
- Modify: `gui/forms/client_transport_cotation.py`
- Test: `tests/test_client_transport_cotation.py`

- [ ] **Step 1: Write the failing regression test for initial rows**

```python
def test_populate_initial_rows_uses_segment_distance(monkeypatch):
    from gui.forms.client_transport_cotation import ClientTransportCotation

    fake = ClientTransportCotation.__new__(ClientTransportCotation)
    fake.client = {
        "ville_depart": "Antananarivo(1 jours)",
        "ville_arrivee": "Antsirabe(2 jours), Ranomafana(2 jours)",
        "itineraire_circuit": "",
    }
    fake._rows = []

    monkeypatch.setattr(
        "gui.forms.client_transport_cotation.load_client_transport_cotation",
        lambda client: [],
    )
    monkeypatch.setattr(
        "gui.forms.client_transport_cotation.get_km_mada_km_for_repere",
        lambda city: {"antananarivo": 0, "antsirabe": 169, "ranomafana": 235}.get(city.lower(), 0),
    )

    fake._populate_initial_rows()

    assert fake._rows[0]["km"] == "169"
    assert fake._rows[1]["km"] == "66"
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `./.env/bin/pytest tests/test_client_transport_cotation.py -k populate_initial_rows -v`

Expected: FAIL because `_populate_initial_rows()` still uses only arrival KM.

- [ ] **Step 3: Implement the minimal runtime fix**

```python
for depart, arrivee in segments:
    km_val = _compute_segment_km(depart, arrivee)
    km_str = str(int(km_val)) if km_val else ""
    self._rows.append(
        _make_row(
            depart=depart,
            arrivee=arrivee,
            km_distance=km_str,
            km=km_str,
        )
    )
```

- [ ] **Step 4: Update the edit dialog arrival/departure callback**

```python
def _on_route_change(*_):
    km_val = _compute_segment_km(v_depart.get().strip(), v_arrivee.get().strip())
    v_km.set(str(int(km_val)) if km_val else "")
    _update_preview()


v_depart.trace_add("write", lambda *a: _on_route_change())
v_arrivee.trace_add("write", lambda *a: _on_route_change())
```

- [ ] **Step 5: Run the client transport tests to verify they pass**

Run: `./.env/bin/pytest tests/test_client_transport_cotation.py -v`

Expected: PASS

- [ ] **Step 6: Commit**

```bash
git add gui/forms/client_transport_cotation.py tests/test_client_transport_cotation.py
git commit -m "Wire normalized segment distances into client transport UI"
```

### Task 4: Ajouter le nettoyage des données existantes `INFOS_CLIENTS`

**Files:**
- Modify: `utils/excel_handler.py`
- Modify: `tests/test_excel_handler.py`
- Test: `tests/test_excel_handler.py`

- [ ] **Step 1: Write the failing migration tests**

```python
def test_normalize_client_route_text_keeps_multi_city_format():
    assert normalize_client_route_text(
        "Antsirabe(2 jours), Ranohira (Isalo)(2 jours)), Tuler(1 jours)"
    ) == "Antsirabe, Ranohira, Toliary"


def test_normalize_existing_client_route_data_updates_infos_clients_sheet(tmp_path, monkeypatch):
    workbook = tmp_path / "clients.xlsx"
    # Build INFOS_CLIENTS with dirty city fields, run migration, then assert normalized cell values.
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `./.env/bin/pytest tests/test_excel_handler.py -k "normalize_client_route_text or normalize_existing_client_route_data" -v`

Expected: FAIL because the normalization and migration functions do not exist yet.

- [ ] **Step 3: Implement text normalization helpers**

```python
def normalize_client_route_text(value):
    parts = re.split(r"[,;\n]+", str(value or ""))
    cleaned = []
    for part in parts:
        normalized = normalize_km_mada_repere(part)
        if normalized:
            cleaned.append(_format_normalized_city_name(normalized))
    return ", ".join(cleaned)
```

- [ ] **Step 4: Implement workbook migration for `INFOS_CLIENTS`**

```python
def normalize_existing_client_route_data(workbook_path=None):
    target = workbook_path or CLIENT_EXCEL_PATH
    wb = load_workbook(target)
    ws = wb[CLIENT_INFOS_SHEET_NAME]
    header_map = _ensure_headers(ws, [])
    for header in ("Ville Départ", "Ville Arrivée", "Itinéraire Circuit"):
        col = header_map.get(header)
        if not col:
            continue
        for row in range(2, ws.max_row + 1):
            current = ws.cell(row=row, column=col).value
            cleaned = normalize_client_route_text(current)
            if cleaned and cleaned != str(current).strip():
                ws.cell(row=row, column=col, value=cleaned)
                logger.info("Normalized %s row %s: %r -> %r", header, row, current, cleaned)
    wb.save(target)
```

- [ ] **Step 5: Run the tests to verify they pass**

Run: `./.env/bin/pytest tests/test_excel_handler.py -k "normalize_client_route_text or normalize_existing_client_route_data" -v`

Expected: PASS

- [ ] **Step 6: Commit**

```bash
git add utils/excel_handler.py tests/test_excel_handler.py
git commit -m "Add INFOS_CLIENTS route normalization migration"
```

### Task 5: Vérification intégrée et exécution de la migration réelle

**Files:**
- Modify: `utils/excel_handler.py`
- Test: `tests/test_excel_handler.py`
- Test: `tests/test_client_transport_cotation.py`

- [ ] **Step 1: Add a targeted smoke test for a real-world alias chain**

```python
def test_real_world_city_alias_chain():
    assert normalize_client_route_text(
        "Antananarivo(1 jours), Antsirabe(2 jours), Ranohira (Isalo)(2 jours)), Tuler(1 jours)"
    ) == "Antananarivo, Antsirabe, Ranohira, Toliary"
```

- [ ] **Step 2: Run the targeted tests**

Run: `./.env/bin/pytest tests/test_excel_handler.py tests/test_client_transport_cotation.py -v`

Expected: PASS

- [ ] **Step 3: Run the migration against the project workbook**

Run: `./.env/bin/python -c "from utils.excel_handler import normalize_existing_client_route_data; print(normalize_existing_client_route_data())"`

Expected: command exits `0` and logs/report indicates the number of normalized cells.

- [ ] **Step 4: Run the broader regression suite**

Run: `./.env/bin/pytest tests/test_excel_handler.py tests/test_models.py tests/test_validators.py tests/test_invoicing.py -v`

Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add utils/excel_handler.py tests/test_excel_handler.py tests/test_client_transport_cotation.py data.xlsx
git commit -m "Normalize transport route data and verify client distance flow"
```
