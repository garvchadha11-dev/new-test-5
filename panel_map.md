# FTA Excise Portal — Verified Panel Map

Captured live from the portal on 2026-04-28. **Note: view prefixes (`__xmlviewNN--`) change every session** — do NOT hardcode them. Use the discovery method (see `discovery_method.md`).

## Findings summary

- All 18 declaration panels use `sap.m.SearchField` for the month search and `sap.m.ComboBox` for status.
- 6 declarations **share a single panel** at `#/ExciseDeclarationList` (`ExciseList_myDecl_Table`): EX203_ML, EX203A, EX203D, EX202A_Enter, EX202A_Production, EX203F. They differ only by which tile was clicked → the underlying form-type filter and status-item options change accordingly.
- Most panels have **no visible Go button** — filtering applies automatically via `fireSearch` + `fireChange`/`fireSelectionChange` on the SAP controls.
- **Status items lazy-load** — most comboboxes show only `["All"]` or `["Select All"]` until first opened. Need `cb.open()` + wait + read items.

## Per-panel control IDs (this session — for reference only)

| Panel | URL hash | Table id | Search control | Status control | Status items observed |
|---|---|---|---|---|---|
| EX201 | `#/201XList` | `__xmlview23--_201X_List_table-listUl` | `_201X_search_searchField` | `_201X_Status_combobox` | Draft, Approved, Cancelled, All |
| EX202B | `#/202BList` | `__xmlview29--_202B_List_table-listUl` | `202B_Search` | `_202B_Status_combobox` | All (lazy) |
| EX203_ML | `#/ExciseDeclarationList` | `__xmlview34--ExciseList_myDecl_Table-listUl` | `ExciseList_myDeclSearch_searchField` | `ExciseList_myDeclStatus_combobox` | All, Approved |
| EX203A | `#/ExciseDeclarationList` | (shared with EX203_ML) | (shared) | (shared) | All, Draft |
| EX203D | `#/ExciseDeclarationList` | (shared) | (shared) | (shared) | All (lazy) |
| EX203H | `#/203HList` | `__xmlview40--_203H_List_table-listUl` | `_203H_List_table_searchField` | `_203H_List_table_combobox` | (lazy) |
| EX202A_Release | `#/202RList` | `__xmlview45--_202R_DeclarationList_table-listUl` | `_202R_Status_searchbar` ⚠️ | `_202R_Status_combobox` | All (lazy) |
| EX202A_Consumption | `#/202SList` | `__xmlview50--_202S_Declaration_List_table-listUl` | `202S_Search` | `_202S_Status_combobox` | All (lazy) |
| EX202A_Enter | `#/ExciseDeclarationList` | (shared) | (shared) | (shared) | All, Approved, Draft, Rejected by Warehouse Keeper |
| EX202A_Transfer | `#/List` | `__xmlview55--202UExciseList_myDecl_Table-listUl` | `_202U_myDecSearch_searchField` | `_202U_myDecStatus_combobox` | All (lazy) |
| EX202A_Export | `#/202VList` | `__xmlview61--_202V_Declaration_List_table-listUl` | `202V_Search` | `_202V_Status_combobox` | All (lazy) |
| EX202A_Import | `#/202WList` | `__xmlview66--_202W_Declaration_List_table-listUl` | `202W_Search` | `_202W_Status_combobox` | Select All (lazy) |
| EX202A_Production | `#/ExciseDeclarationList` | (shared) | (shared) | (shared) | All (lazy) |
| EX203B | `#/203BList` | `__xmlview71--_203B_Declaration_List_table-listUl` | `_203B_Declaration_ListSearch_searchField` | `_203B_Declaration_ListStatus_combobox` | Select All (lazy) |
| EX203C | `#/203CList` | `__xmlview77--203CExciseList_myDecl_Table-listUl` | `_203C_myDecSearch_searchField` | `_203C_myDecStatus_combobox` | All, Approved, Draft, Awaiting Purchaser Approval, Rejected by Purchaser |
| EX203F | `#/ExciseDeclarationList` | (shared) | (shared) | (shared) | All (lazy) |
| EX204 | `#/204XList` | `__xmlview83--_204X_Declaration_List_table-listUl` | `204X_Search` | `_204X_Status_combobox` | Select All (lazy) |
| EX203G | `#/203GList` | `__xmlview88--_203G_Declaration_List_table-listUl` | `203G_Search` | `_203G_Status_combobox` | All (lazy) |

⚠️ EX202A_Release names its search field `_Status_searchbar` (not `_Search` or `_searchField`) — proves there's no consistent naming convention. **Only the SAP API + dynamic discovery is reliable.**

## Naming conventions observed (so future-you doesn't trust patterns)

Search field id suffixes seen, all on the same portal:
- `_search_searchField` (EX201)
- `_Search` (EX202B, 202S, 202V, 202W, 203G, 204X)
- `_searchField` (EX203H)
- `_DeclarationListSearch_searchField` (EX203B)
- `Search_searchField` (myDecl variants — EX203_ML, EX203C)
- `_Status_searchbar` (EX202A_Release — labeled as Status but is the search field!)
- `_myDecSearch_searchField` (EX202A_Transfer, EX203C)

Don't pattern-match. Use SAP discovery.
