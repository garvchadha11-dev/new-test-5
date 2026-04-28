# Discovery method (verified across 18 panels)

This replaces the broken hardcoded `SEARCH_IDS` / `STATUS_IDS` / view-prefix maps in v5.

## Step 1 — find the active panel

```js
function findActivePanel() {
  function visible(el) {
    if (!el) return false;
    var r = el.getBoundingClientRect();
    if (r.width <= 0 || r.height <= 0) return false;
    // SAP keeps stale pages mounted but pushes them off-viewport
    if (r.right < 0 || r.left > window.innerWidth) return false;
    return true;
  }
  var allTables = Array.from(document.querySelectorAll('table[id*="-listUl"]')).filter(visible);
  for (var i = 0; i < allTables.length; i++) {
    var t = allTables[i];
    var dash = t.id.indexOf('--');
    if (dash < 0) continue;
    var prefix = t.id.substring(0, dash + 2);
    var search = Array.from(document.querySelectorAll('input[type="search"]'))
      .filter(function(el){ return el.id.indexOf(prefix) === 0 && visible(el); })[0];
    var statusInner = Array.from(document.querySelectorAll('input[id$="_combobox-inner"]'))
      .filter(function(el){ return el.id.indexOf(prefix) === 0 && visible(el); })[0];
    if (search && statusInner) {
      return {
        prefix: prefix,
        table: t.id,
        searchControlId: search.id.replace(/-I$/, ''),
        statusControlId: statusInner.id.replace(/-inner$/, '')
      };
    }
  }
  return null;
}
```

The "search input AND status combobox both in same view prefix, both on-viewport" predicate uniquely identifies the active panel — verified on every one of the 18 declarations.

## Step 2 — apply the month filter via SAP API

```js
var sf = sap.ui.getCore().byId(panel.searchControlId);
sf.setValue('march 2026');
sf.fireSearch({ query: 'march 2026' });
```

Confirmed working — every panel's search field is `sap.m.SearchField` with `setValue` + `fireSearch`.

## Step 3 — apply the status filter via SAP API

```js
var cb = sap.ui.getCore().byId(panel.statusControlId);
// Items lazy-load — open the dropdown, wait, then operate
cb.open();
// (small wait — 500ms usually enough)
var match = cb.getItems().find(function(it) { return it.getText() === 'Approved'; });
if (match) {
  cb.setSelectedItem(match);
  cb.setValue(match.getText());
  cb.fireChange({ value: match.getText(), itemPressed: true });
  cb.fireSelectionChange({ selectedItem: match });
}
cb.close();
```

If `Approved` isn't in the items list, fall back to `Approved by Warehouse Keeper` (only on EX202A_* panels), or just leave on the default and let the search filter do the work.

## Step 4 — fire Go (only if visible)

```js
var goBtn = Array.from(document.querySelectorAll('button')).find(function(b) {
  if (b.id.indexOf(panel.prefix) !== 0) return false;
  var bdi = b.querySelector('bdi');
  return bdi && bdi.textContent.trim() === 'Go' && b.getBoundingClientRect().width > 0;
});
if (goBtn) sap.ui.getCore().byId(goBtn.id).firePress();
```

Most panels don't have a visible Go button — `fireSearch` + `fireChange` are enough to refresh the table.
