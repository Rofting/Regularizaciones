"""Generación auditable de cartas Word para un expediente ya conciliado."""

from __future__ import annotations

import hashlib
import json
import os
import re
import sqlite3
import uuid
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Callable

from carta_writer import generar_carta
from excel_export_service import ExportBlockedError, calculate_case_input_hash
from excel_profiles import ExcelProfile, load_profile
from letter_settings import LetterIdentity, load_community_letter_identity


class LetterGenerationBlockedError(ValueError):
    """El expediente no reúne las garantías para emitir cartas."""


@dataclass(frozen=True)
class LetterBatchResult:
    id_letter_run: int
    output_path: Path
    generated_count: int
    failures: tuple[str, ...]


ProgressCallback = Callable[[str, dict[str, Any]], None]


def _emit(progress: ProgressCallback | None, stage: str, **payload: Any) -> None:
    if progress is not None:
        progress(stage, payload)


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        while chunk := handle.read(64 * 1024):
            digest.update(chunk)
    return digest.hexdigest()


def _safe_file_part(value: object, *, fallback: str) -> str:
    cleaned = re.sub(r"[^\w\s-]", "", str(value or ""), flags=re.UNICODE).strip()
    cleaned = re.sub(r"\s+", "_", cleaned)
    return cleaned.upper()[:50] or fallback


def _case_context(connection: sqlite3.Connection, id_case: int) -> sqlite3.Row:
    row = connection.execute(
        """SELECT c.id_case,c.id_comunidad,c.id_periodo,c.fecha_inicio,c.fecha_fin,
                  c.estado,co.codigo,co.nombre AS community_name,
                  per.nombre AS period_name
           FROM regularization_cases c
           JOIN comunidades co ON co.id_comunidad=c.id_comunidad
           LEFT JOIN periodos per ON per.id_periodo=c.id_periodo
           WHERE c.id_case=?""",
        (id_case,),
    ).fetchone()
    if row is None:
        raise LetterGenerationBlockedError("El expediente no existe")
    if row["id_periodo"] is None or row["period_name"] is None:
        raise LetterGenerationBlockedError("El expediente no tiene un período ligado")
    if row["estado"] not in ("calculated", "reconciled"):
        raise LetterGenerationBlockedError("El expediente aún no está calculado y conciliado")
    period = connection.execute(
        """SELECT id_comunidad,fecha_inicio,fecha_fin FROM periodos WHERE id_periodo=?""",
        (row["id_periodo"],),
    ).fetchone()
    if (
        period is None or int(period["id_comunidad"]) != int(row["id_comunidad"])
        or period["fecha_inicio"] != row["fecha_inicio"]
        or period["fecha_fin"] != row["fecha_fin"]
    ):
        raise LetterGenerationBlockedError("El período ligado no coincide con el expediente")
    issues = connection.execute(
        "SELECT COUNT(*) FROM review_issues WHERE id_case=? AND status='open'", (id_case,)
    ).fetchone()[0]
    if issues:
        raise LetterGenerationBlockedError(
            f"Hay {issues} incidencia(s) abierta(s) antes de generar las cartas"
        )
    return row


def _validated_export(
    connection: sqlite3.Connection, case: sqlite3.Row, project_root: Path
) -> tuple[sqlite3.Row, ExcelProfile]:
    export = connection.execute(
        """SELECT e.id_export_run,e.id_periodo,e.input_sha256,e.template_sha256,e.status,
                  t.profile_key,t.profile_version
           FROM excel_export_runs e
           JOIN excel_template_profiles t ON t.id_template_profile=e.id_template_profile
           WHERE e.id_case=? ORDER BY e.created_at DESC,e.id_export_run DESC LIMIT 1""",
        (case["id_case"],),
    ).fetchone()
    if export is None or export["status"] != "validated":
        raise LetterGenerationBlockedError("Falta un Excel oficial validado para el expediente")
    if int(export["id_periodo"]) != int(case["id_periodo"]):
        raise LetterGenerationBlockedError("El Excel validado pertenece a otro período")
    try:
        profile = load_profile(export["profile_key"], project_root)
    except (LookupError, ValueError) as error:
        raise LetterGenerationBlockedError(f"No se puede cargar el perfil del Excel: {error}") from error
    if profile.version != export["profile_version"]:
        raise LetterGenerationBlockedError("La versión del perfil no coincide con el Excel validado")
    if profile.community_code != str(case["codigo"]):
        raise LetterGenerationBlockedError("El perfil del Excel no corresponde a la comunidad")
    try:
        current_hash = calculate_case_input_hash(
            connection, id_case=int(case["id_case"]), project_root=project_root, profile=profile
        )
    except ExportBlockedError as error:
        raise LetterGenerationBlockedError(str(error)) from error
    if export["input_sha256"] != current_hash:
        raise LetterGenerationBlockedError(
            "Las entradas cambiaron desde el Excel validado; debe regenerar Excel antes de las cartas"
        )
    return export, profile


def _active_results(
    connection: sqlite3.Connection, case: sqlite3.Row, profile: ExcelProfile
) -> tuple[tuple[str, ...], list[sqlite3.Row]]:
    profile_keys = tuple(concept.key for concept in profile.concepts)
    if not profile_keys:
        raise LetterGenerationBlockedError("El perfil no declara conceptos para las cartas")
    placeholders = ",".join("?" for _ in profile_keys)
    rows = connection.execute(
        f"""SELECT r.id_propietario,r.concept_key,r.consumption,r.consumption_unit,
                   r.billed_cents,r.actual_cents,r.difference_cents,
                   rc.label,rc.service,rc.unit,rc.display_order,
                   p.codigo_vivienda,p.nombre_propietario,p.coeficiente
            FROM owner_concept_results r
            JOIN propietarios p ON p.id_propietario=r.id_propietario
            JOIN regularization_concepts rc ON rc.concept_key=r.concept_key
            WHERE r.id_periodo=? AND p.id_comunidad=? AND p.activo=1
              AND r.concept_key IN ({placeholders})
            ORDER BY p.id_propietario,rc.display_order,r.concept_key""",
        (case["id_periodo"], case["id_comunidad"], *profile_keys),
    ).fetchall()
    active_keys = tuple(
        key for key in profile_keys if any(row["concept_key"] == key for row in rows)
    )
    if not active_keys or not rows:
        raise LetterGenerationBlockedError("No hay resultados canónicos para las cartas")
    owner_ids = [
        int(row[0]) for row in connection.execute(
            "SELECT id_propietario FROM propietarios WHERE id_comunidad=? AND activo=1 ORDER BY id_propietario",
            (case["id_comunidad"],),
        ).fetchall()
    ]
    result_keys = {(int(row["id_propietario"]), row["concept_key"]) for row in rows}
    missing = [
        f"propietario {owner_id}/{key}" for owner_id in owner_ids for key in active_keys
        if (owner_id, key) not in result_keys
    ]
    if missing:
        raise LetterGenerationBlockedError(
            "Faltan resultados canónicos para " + ", ".join(missing[:3])
        )
    reconciliation = {
        row["concept_key"]: row["status"]
        for row in connection.execute(
            f"""SELECT concept_key,status FROM reconciliations
                WHERE id_periodo=? AND concept_key IN ({','.join('?' for _ in active_keys)})""",
            (case["id_periodo"], *active_keys),
        ).fetchall()
    }
    unbalanced = [key for key in active_keys if reconciliation.get(key) != "cuadrado"]
    if unbalanced:
        raise LetterGenerationBlockedError(
            "El reparto no está conciliado para: " + ", ".join(unbalanced)
        )
    return active_keys, rows


def _consumption_graphs(
    connection: sqlite3.Connection,
    case: sqlite3.Row,
    owners: list[dict[str, Any]],
) -> dict[int, dict[str, Any]]:
    """Prepara sólo la comparativa vecinal e histórico, nunca por períodos ajenos."""
    readings = connection.execute(
        """SELECT l.id_propietario,l.fecha_lectura,l.valor_acumulado
           FROM lecturas_vecino l JOIN propietarios p ON p.id_propietario=l.id_propietario
           WHERE p.id_comunidad=? AND p.activo=1 AND l.id_periodo=? AND l.tipo='ACS'
             AND l.fecha_lectura IN (?,?)
           ORDER BY l.id_propietario,l.fecha_lectura""",
        (case["id_comunidad"], case["id_periodo"], case["fecha_inicio"], case["fecha_fin"]),
    ).fetchall()
    values: dict[int, dict[str, float]] = {}
    for reading in readings:
        values.setdefault(int(reading["id_propietario"]), {})[reading["fecha_lectura"]] = float(
            reading["valor_acumulado"]
        )
    consumption = {
        owner_id: max(0.0, value[case["fecha_fin"]] - value[case["fecha_inicio"]])
        for owner_id, value in values.items()
        if case["fecha_inicio"] in value and case["fecha_fin"] in value
    }
    neighbors = list(consumption.values())
    graphs: dict[int, dict[str, Any]] = {}
    for owner in owners:
        owner_id = int(owner["id_propietario"])
        historical = connection.execute(
            """SELECT per.nombre,ini.valor_acumulado AS initial_value,fin.valor_acumulado AS final_value
               FROM periodos per
               JOIN lecturas_vecino ini ON ini.id_periodo=per.id_periodo
                   AND ini.id_propietario=? AND ini.tipo='ACS' AND ini.fecha_lectura=per.fecha_inicio
               JOIN lecturas_vecino fin ON fin.id_periodo=per.id_periodo
                   AND fin.id_propietario=? AND fin.tipo='ACS' AND fin.fecha_lectura=per.fecha_fin
               WHERE per.id_comunidad=? AND per.id_periodo<>? ORDER BY per.fecha_inicio""",
            (owner_id, owner_id, case["id_comunidad"], case["id_periodo"]),
        ).fetchall()
        graphs[owner_id] = {
            "owner_consumption": consumption.get(owner_id, 0.0),
            "neighbor_consumptions": neighbors,
            "history": [
                (row["nombre"], max(0.0, float(row["final_value"]) - float(row["initial_value"])))
                for row in historical
            ],
            "unit": "m³",
        }
    return graphs


def _letter_input_hash(
    export: sqlite3.Row, identity: LetterIdentity, rows: list[sqlite3.Row]
) -> str:
    canonical_rows = [
        [
            int(row["id_propietario"]), row["concept_key"], row["consumption"],
            int(row["billed_cents"]), int(row["actual_cents"]), int(row["difference_cents"]),
        ]
        for row in rows
    ]
    identity_payload = asdict(identity)
    identity_payload["logo_path"] = str(identity.logo_path) if identity.logo_path else None
    payload = json.dumps(
        {"excel_input": export["input_sha256"], "rows": canonical_rows, "identity": identity_payload},
        ensure_ascii=False, sort_keys=True, separators=(",", ":"),
    )
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def _write_run_status(
    connection: sqlite3.Connection, run_id: int, status: str, error: str | None = None
) -> None:
    connection.execute(
        """UPDATE letter_generation_runs SET status=?,error_message=?,updated_at=datetime('now'),
               completed_at=CASE WHEN ? THEN datetime('now') ELSE completed_at END
           WHERE id_letter_run=?""",
        (status, error, 1 if status in ("completed", "incomplete", "failed") else 0, run_id),
    )
    connection.commit()


def _is_safe_generated_path(raw_path: object, output_directory: Path) -> bool:
    """Acepta únicamente archivos existentes bajo el directorio de esta salida."""
    if not raw_path:
        return False
    root = output_directory.resolve()
    candidate = Path(str(raw_path)).resolve()
    try:
        candidate.relative_to(root)
    except ValueError:
        return False
    return candidate.is_file()


def _reusable_completed_run(
    connection: sqlite3.Connection,
    *,
    case: sqlite3.Row,
    export: sqlite3.Row,
    input_hash: str,
    template_hash: str,
    output_directory: Path,
    expected_owner_ids: set[int],
) -> LetterBatchResult | None:
    """Devuelve sólo un lote completo cuya auditoría y archivos siguen íntegros."""
    runs = connection.execute(
        """SELECT id_letter_run,output_path FROM letter_generation_runs
           WHERE id_case=? AND id_periodo=? AND id_export_run=?
             AND input_sha256=? AND template_sha256=? AND status='completed'
           ORDER BY completed_at DESC,id_letter_run DESC""",
        (
            case["id_case"], case["id_periodo"], export["id_export_run"],
            input_hash, template_hash,
        ),
    ).fetchall()
    period_directory = output_directory.resolve()
    for run in runs:
        try:
            run_directory = Path(run["output_path"]).resolve()
            run_directory.relative_to(period_directory)
            if run_directory == period_directory:
                continue
        except (OSError, TypeError):
            continue
        rows = connection.execute(
            """SELECT id_propietario,status,output_path FROM generated_letters
               WHERE id_letter_run=? ORDER BY id_propietario""",
            (run["id_letter_run"],),
        ).fetchall()
        if {int(row["id_propietario"]) for row in rows} != expected_owner_ids:
            continue
        if len(rows) != len(expected_owner_ids):
            continue
        if all(
            row["status"] == "generated"
            and _is_safe_generated_path(row["output_path"], run_directory)
            for row in rows
        ):
            return LetterBatchResult(
                int(run["id_letter_run"]), run_directory, len(rows), ()
            )
    return None


def _temporary_destination(destination: Path) -> Path:
    """Reserva un nombre hermano no visible como carta final."""
    return destination.with_name(f".{destination.stem}.{uuid.uuid4().hex}.tmp")


def _mark_generated(
    connection: sqlite3.Connection, generated_row: int, destination: Path
) -> None:
    """Confirma que el archivo publicado tiene una auditoría coherente."""
    connection.execute(
        """UPDATE generated_letters SET status='generated',output_path=?,error_message=NULL,
               updated_at=datetime('now') WHERE id_generated_letter=?""",
        (str(destination), generated_row),
    )
    connection.commit()


def generate_case_letters(
    database_path: str | Path,
    *,
    id_case: int,
    project_root: Path,
    progress: ProgressCallback | None = None,
) -> LetterBatchResult:
    """Genera cartas una a una, auditando el resultado de cada propietario."""
    root = Path(project_root).resolve()
    _emit(progress, "validate_case", id_case=id_case)
    connection = sqlite3.connect(str(database_path))
    connection.row_factory = sqlite3.Row
    try:
        case = _case_context(connection, id_case)
        export, profile = _validated_export(connection, case, root)
        active_keys, rows = _active_results(connection, case, profile)
        identity = load_community_letter_identity(root, str(case["codigo"]))
        template = root / "plantillas" / "Plantilla_Cartas.docx"
        if not template.is_file():
            raise LetterGenerationBlockedError(f"No existe la plantilla de cartas: {template}")
        output_path = root / "salidas" / "cartas" / _safe_file_part(case["codigo"], fallback="COMUNIDAD") / _safe_file_part(
            case["period_name"], fallback="PERIODO"
        )
        input_hash = _letter_input_hash(export, identity, rows)
        template_hash = _sha256(template)
        by_owner: dict[int, list[sqlite3.Row]] = {}
        owner_data: dict[int, dict[str, Any]] = {}
        for row in rows:
            owner_id = int(row["id_propietario"])
            by_owner.setdefault(owner_id, []).append(row)
            owner_data.setdefault(owner_id, {
                "id_propietario": owner_id,
                "nombre": row["nombre_propietario"],
                "vivienda": row["codigo_vivienda"],
                "coeficiente": row["coeficiente"],
            })
        graphs = _consumption_graphs(connection, case, list(owner_data.values()))

        reusable = _reusable_completed_run(
            connection,
            case=case,
            export=export,
            input_hash=input_hash,
            template_hash=template_hash,
            output_directory=output_path,
            expected_owner_ids=set(by_owner),
        )
        if reusable is not None:
            _emit(progress, "reused", id_letter_run=reusable.id_letter_run)
            return reusable

        _emit(progress, "prepare_output", output_path=str(output_path))
        output_path.mkdir(parents=True, exist_ok=True)
        cursor = connection.execute(
            """INSERT INTO letter_generation_runs
               (id_case,id_periodo,id_export_run,input_sha256,template_sha256,output_path,status)
               VALUES (?,?,?,?,?,?, 'running')""",
            (id_case, case["id_periodo"], export["id_export_run"], input_hash, template_hash, str(output_path)),
        )
        run_id = int(cursor.lastrowid)
        output_path = output_path / f"lote_{run_id}"
        connection.execute(
            """UPDATE letter_generation_runs SET output_path=?,updated_at=datetime('now')
               WHERE id_letter_run=?""",
            (str(output_path), run_id),
        )
        connection.commit()
        try:
            output_path.mkdir(parents=False, exist_ok=False)
        except Exception as error:
            _write_run_status(connection, run_id, "failed", f"No se pudo crear salida: {error}")
            raise

        failures: list[str] = []
        generated = 0
        for owner_id in sorted(by_owner):
            owner = owner_data[owner_id]
            concepts = [
                {
                    "key": row["concept_key"], "label": row["label"], "service": row["service"],
                    "unit": row["unit"], "consumption": row["consumption"],
                    "importe_cobrado": int(row["billed_cents"]) / 100,
                    "importe_real": int(row["actual_cents"]) / 100,
                    "diferencia": int(row["difference_cents"]) / 100,
                }
                for row in by_owner[owner_id] if row["concept_key"] in active_keys
            ]
            total = {
                "cobrado": round(sum(item["importe_cobrado"] for item in concepts), 2),
                "real": round(sum(item["importe_real"] for item in concepts), 2),
                "diferencia": round(sum(item["diferencia"] for item in concepts), 2),
            }
            filename = "CARTA_{}_{}.docx".format(
                _safe_file_part(owner["vivienda"], fallback=f"P{owner_id}"),
                _safe_file_part(owner["nombre"], fallback=f"P{owner_id}"),
            )
            destination = output_path / filename
            temporary_destination = _temporary_destination(destination)
            generated_row = connection.execute(
                """INSERT INTO generated_letters
                   (id_letter_run,id_propietario,input_sha256,template_sha256,status)
                   VALUES (?,?,?,?, 'pending')""",
                (run_id, owner_id, input_hash, template_hash),
            ).lastrowid
            connection.commit()
            _emit(progress, "generate_letter", id_propietario=owner_id, total=len(by_owner))
            letter_data = {
                "vecino": {"nombre": owner["nombre"], "vivienda": owner["vivienda"], "coeficiente": owner["coeficiente"]},
                "periodo": {"nombre": case["period_name"], "fecha_inicio": case["fecha_inicio"], "fecha_fin": case["fecha_fin"]},
                "comunidad": case["community_name"], "conceptos": concepts, "total": total,
                "medias": {"n_vecinos": len(by_owner)}, "consumo_grafica": graphs.get(owner_id, {}),
                "identity": {
                    "office_name": identity.office_name, "footer": identity.footer,
                    "signature": identity.signature, "city": identity.city,
                    "logo_path": identity.logo_path,
                },
            }
            published = False
            try:
                generar_carta(letter_data, str(template), str(temporary_destination))
                os.replace(temporary_destination, destination)
                published = True
                _mark_generated(connection, generated_row, destination)
                generated += 1
            except Exception as error:
                if published:
                    try:
                        destination.unlink(missing_ok=True)
                    except OSError:
                        pass
                try:
                    temporary_destination.unlink(missing_ok=True)
                except OSError:
                    pass
                message = f"{owner['nombre']}: {type(error).__name__}: {error}"
                failures.append(message)
                try:
                    connection.rollback()
                    connection.execute(
                        """UPDATE generated_letters SET status='failed',output_path=NULL,error_message=?,
                               updated_at=datetime('now') WHERE id_generated_letter=?""",
                        (message, generated_row),
                    )
                    connection.commit()
                except sqlite3.Error:
                    try:
                        connection.rollback()
                    except sqlite3.Error:
                        pass

        status = "completed" if not failures else "incomplete"
        _write_run_status(connection, run_id, status, "; ".join(failures) if failures else None)
        _emit(progress, status, id_letter_run=run_id, generated_count=generated, failures=len(failures))
        return LetterBatchResult(run_id, output_path, generated, tuple(failures))
    finally:
        connection.close()
