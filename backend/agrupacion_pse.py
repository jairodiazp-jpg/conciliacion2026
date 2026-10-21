from typing import List, Dict, Callable, Tuple

VALUE_TOL = 0.01


def _to_float(val) -> float:
    if val is None or val == "":
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    # normalize thousand separators and decimals
    s = s.replace(".", "").replace(",", ".") if s.count(',') == 1 and s.count('.') > 1 else s.replace(',', '')
    try:
        return float(s)
    except Exception:
        try:
            return float(s.replace('\u00A0', ''))
        except Exception:
            return 0.0


def _next_group_id(existing_max: int) -> Tuple[str, int]:
    n = existing_max + 1
    return f"GRP-{n:04d}", n


def conciliacion_por_agrupacion(ws, pse_dataset: List[Dict], col_map: Dict[str,int],
                                find_max_grp_number: Callable[[], int],
                                marca_pse_predicate: Callable[[object, int], bool]):
    """
    Ejecuta la fase de 'Conciliación por agrupación PSE' sobre la hoja `ws`.

    - `ws`: openpyxl worksheet
    - `pse_dataset`: lista de dicts con al menos la clave 'valor'
    - `col_map`: {'valor':int,'estado':int,'id_grupo':int,'comentario':int,'descripcion':int}
    - `find_max_grp_number`: función que devuelve el max número GRP ya existente (int)
    - `marca_pse_predicate`: funcion(ws, row_idx) -> True si la fila está marcada como PSE
    """
    max_grp = find_max_grp_number()
    used_rows = set()
    # asumimos encabezado en la fila 1
    for i in range(2, ws.max_row + 1):
        if i in used_rows:
            continue
        estado_cell = ws.cell(row=i, column=col_map['estado'])
        estado_val = (estado_cell.value or "").strip() if estado_cell.value else ""
        if estado_val in ("Cruzado", "Posible cruce"):
            continue
        # sólo filas marcadas PSE como potencial total
        if not marca_pse_predicate(ws, i):
            continue
        total_val = _to_float(ws.cell(row=i, column=col_map['valor']).value)
        if abs(total_val) < VALUE_TOL:
            continue
        # buscar desgloses contiguos debajo
        breakdown_rows = []
        j = i + 1
        while j <= ws.max_row:
            estado_next = (ws.cell(row=j, column=col_map['estado']).value or "").strip() if ws.cell(row=j, column=col_map['estado']).value else ""
            if estado_next in ("Cruzado", "Posible cruce"):
                break
            val = _to_float(ws.cell(row=j, column=col_map['valor']).value)
            if abs(val) < VALUE_TOL:
                break
            breakdown_rows.append((j, val))
            if sum(v for (_, v) in breakdown_rows) >= total_val - VALUE_TOL:
                break
            j += 1
        if not breakdown_rows:
            continue
        sum_break = sum(v for (_, v) in breakdown_rows)
        if abs(sum_break - total_val) > VALUE_TOL:
            # no coincide la suma
            ws.cell(row=i, column=col_map['comentario']).value = f"Suma desgloses {sum_break:.2f} != total {total_val:.2f}"
            ws.cell(row=i, column=col_map['estado']).value = "Parcial"
            continue
        # validar existencia en pse_dataset
        missing = []
        for (_, val) in breakdown_rows:
            found = any(abs(_to_float(p.get('valor') or p.get('monto') or 0) - val) <= VALUE_TOL for p in pse_dataset)
            if not found:
                missing.append(val)
        if missing:
            ws.cell(row=i, column=col_map['comentario']).value = f"Faltan desgloses en PSE: {missing}"
            ws.cell(row=i, column=col_map['estado']).value = "Parcial"
            continue
        # crear grupo
        new_grp_str, max_grp = _next_group_id(max_grp)
        # asignar al total
        ws.cell(row=i, column=col_map['id_grupo']).value = new_grp_str
        ws.cell(row=i, column=col_map['estado']).value = "Conciliado por agrupación"
        ws.cell(row=i, column=col_map['comentario']).value = f"Agrupación PSE {new_grp_str} - total {total_val:.2f}"
        conflict = False
        for (ridx, val) in breakdown_rows:
            if ws.cell(row=ridx, column=col_map['id_grupo']).value:
                # conflicto: revertir
                conflict = True
                break
            ws.cell(row=ridx, column=col_map['id_grupo']).value = new_grp_str
            ws.cell(row=ridx, column=col_map['estado']).value = "Conciliado por agrupación"
            ws.cell(row=ridx, column=col_map['comentario']).value = f"Agrupación PSE {new_grp_str} - valor {val:.2f}"
            used_rows.add(ridx)
        if conflict:
            # revertir asignaciones parciales
            ws.cell(row=i, column=col_map['id_grupo']).value = None
            ws.cell(row=i, column=col_map['estado']).value = "Parcial"
            ws.cell(row=i, column=col_map['comentario']).value = "Conflicto: algún desglose ya pertenece a otro grupo"
            # limpiar los que asignamos
            for (ridx, _) in breakdown_rows:
                if ws.cell(row=ridx, column=col_map['id_grupo']).value == new_grp_str:
                    ws.cell(row=ridx, column=col_map['id_grupo']).value = None
                    ws.cell(row=ridx, column=col_map['estado']).value = None
                    ws.cell(row=ridx, column=col_map['comentario']).value = None
            continue
        used_rows.add(i)
