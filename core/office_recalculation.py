"""Recálculo controlado de libros mediante LibreOffice o una doble de pruebas."""

from __future__ import annotations

import os
import posixpath
import re
import shutil
import subprocess
import tempfile
import time
from copy import deepcopy
from pathlib import Path
from typing import Callable, Protocol
from xml.etree import ElementTree as ET
from zipfile import ZipFile


_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def _worksheet_paths(archive: ZipFile) -> dict[str, str]:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {
        item.attrib["Id"]: item.attrib["Target"]
        for item in relationships.findall(f"{{{_PKG_REL}}}Relationship")
    }
    paths = {}
    for sheet in workbook.findall(f"{{{_MAIN}}}sheets/{{{_MAIN}}}sheet"):
        target = targets[sheet.attrib[f"{{{_REL}}}id"]]
        paths[sheet.attrib["name"]] = (
            target.lstrip("/") if target.startswith("/")
            else posixpath.normpath(posixpath.join("xl", target))
        )
    return paths


def restore_design_with_calculated_values(original: Path, calculated: Path) -> None:
    """Conserva el OOXML diseñado por la plantilla y copia los cachés de fórmulas.

    LibreOffice recalcula correctamente, pero reescribe fuentes, colores y otros
    detalles de presentación. Los únicos cambios necesarios son los resultados
    guardados en las celdas con fórmula.
    """
    replacement: dict[str, bytes] = {}
    with ZipFile(original) as source, ZipFile(calculated) as result:
        source_sheets = _worksheet_paths(source)
        result_sheets = _worksheet_paths(result)
        shared_strings = []
        if "xl/sharedStrings.xml" in result.namelist():
            string_table = ET.fromstring(result.read("xl/sharedStrings.xml"))
            shared_strings = [
                "".join(node.text or "" for node in item.iter(f"{{{_MAIN}}}t"))
                for item in string_table.findall(f"{{{_MAIN}}}si")
            ]
        for name, source_path in source_sheets.items():
            if name not in result_sheets:
                raise RecalculationError(f"LibreOffice eliminó la hoja {name}")
            source_xml = ET.fromstring(source.read(source_path))
            result_xml = ET.fromstring(result.read(result_sheets[name]))
            result_cells = {
                cell.attrib["r"]: cell
                for cell in result_xml.iter(f"{{{_MAIN}}}c")
                if "r" in cell.attrib
            }
            changed = False
            for cell in source_xml.iter(f"{{{_MAIN}}}c"):
                formula = cell.find(f"{{{_MAIN}}}f")
                if formula is None:
                    continue
                address = cell.attrib.get("r", "")
                computed = result_cells.get(address)
                if computed is None or computed.find(f"{{{_MAIN}}}f") is None:
                    raise RecalculationError(f"Falta la fórmula recalculada {name}!{address}")
                value = computed.find(f"{{{_MAIN}}}v")
                if value is None:
                    raise RecalculationError(f"Falta el resultado recalculado {name}!{address}")
                if computed.attrib.get("t") == "s":
                    try:
                        value = ET.Element(f"{{{_MAIN}}}v")
                        value.text = shared_strings[int(computed.find(f"{{{_MAIN}}}v").text)]
                    except (IndexError, TypeError, ValueError) as error:
                        raise RecalculationError(
                            f"No se puede leer el resultado de {name}!{address}"
                        ) from error
                previous = cell.find(f"{{{_MAIN}}}v")
                if previous is not None:
                    cell.remove(previous)
                cell.insert(list(cell).index(formula) + 1, deepcopy(value))
                if "t" in computed.attrib:
                    cell.attrib["t"] = (
                        "str" if computed.attrib["t"] == "s" else computed.attrib["t"]
                    )
                else:
                    cell.attrib.pop("t", None)
                changed = True
            if changed:
                replacement[source_path] = ET.tostring(source_xml, encoding="utf-8")
        temporary = original.with_name(original.name + ".cached")
        try:
            with ZipFile(temporary, "w") as output:
                for part in source.infolist():
                    output.writestr(part, replacement.get(part.filename, source.read(part.filename)))
        except BaseException:
            temporary.unlink(missing_ok=True)
            raise
    # Windows no permite reemplazar el libro mientras ZipFile lo tiene abierto.
    try:
        os.replace(temporary, calculated)
    finally:
        temporary.unlink(missing_ok=True)


# Versión más antigua con la que se ha verificado el recorrido completo
# (mejora 12: 24.2.7.2 en Ubuntu 24.04 y 26.2.6.3 en Windows).
MINIMUM_LIBREOFFICE_VERSION = (24, 2)


def parse_libreoffice_version(text: str) -> tuple[int, ...] | None:
    """Extrae la versión de la salida de ``soffice --version``."""
    match = re.search(r"LibreOffice\s+(\d+(?:\.\d+)+)", text or "")
    return tuple(int(part) for part in match.group(1).split(".")) if match else None


def version_text(version: tuple[int, ...]) -> str:
    return ".".join(str(part) for part in version)


class RecalculationError(RuntimeError):
    """Indica que el libro no pudo recalcularse con un motor de escritorio."""


class WorkbookRecalculator(Protocol):
    def recalculate(self, workbook_path: Path, work_directory: Path) -> None: ...


class DeterministicRecalculator:
    """Doble inyectable que permite verificar el flujo sin LibreOffice."""

    def __init__(
        self,
        after_recalculate: Callable[[Path], None] | None = None,
    ) -> None:
        self._after_recalculate = after_recalculate

    def recalculate(self, workbook_path: Path, work_directory: Path) -> None:
        if not Path(workbook_path).is_file():
            raise RecalculationError(f"No existe el libro que se quiere recalcular: {workbook_path}")
        Path(work_directory).mkdir(parents=True, exist_ok=True)
        if self._after_recalculate is not None:
            self._after_recalculate(Path(workbook_path))


class LibreOfficeRecalculator:
    """Recalcula una copia XLSX con LibreOffice en modo headless."""

    def __init__(
        self,
        *,
        executable: str | Path | None = None,
        timeout_seconds: int = 90,
    ) -> None:
        if timeout_seconds <= 0:
            raise ValueError("El timeout de LibreOffice debe ser positivo")
        self._executable = Path(executable) if executable is not None else None
        self._timeout_seconds = timeout_seconds

    def _find_executable(self) -> Path:
        def waiting_launcher(path: Path) -> Path:
            # En Windows soffice.exe es un lanzador gráfico: puede devolver el
            # control antes de que termine la conversión y el programa cree
            # erróneamente que no hubo salida. soffice.com es el lanzador de
            # consola equivalente y sí espera al proceso real.
            console = path.with_suffix(".com")
            if os.name == "nt" and console.is_file():
                return console
            return path

        if self._executable is not None:
            if self._executable.is_file():
                return waiting_launcher(self._executable)
            raise RecalculationError(
                "LibreOffice no está disponible en la ruta configurada. "
                "Instálalo o selecciona su ejecutable soffice."
            )
        for command in ("soffice", "libreoffice"):
            located = shutil.which(command)
            if located:
                return waiting_launcher(Path(located))
        candidates: list[Path] = []
        for variable in ("ProgramFiles", "ProgramFiles(x86)"):
            base = os.environ.get(variable)
            if base:
                program = Path(base) / "LibreOffice" / "program"
                candidates.extend((program / "soffice.com", program / "soffice.exe"))
        for candidate in candidates:
            if candidate.is_file():
                return candidate
        raise RecalculationError(
            "LibreOffice no está instalado o no se encuentra soffice. "
            "Instálalo para poder validar el Excel oficial."
        )

    def installed_version(self) -> tuple[int, ...]:
        """Versión del LibreOffice que se usaría para recalcular."""
        executable = self._find_executable()
        with tempfile.TemporaryDirectory(prefix="regularizacion-lo-") as profile_root:
            try:
                completed = subprocess.run(
                    [str(executable), f"-env:UserInstallation={Path(profile_root).resolve().as_uri()}", "--version"],
                    capture_output=True, text=True, timeout=self._timeout_seconds, check=False,
                )
            except (OSError, subprocess.TimeoutExpired) as error:
                raise RecalculationError(f"No se pudo consultar la versión de LibreOffice: {error}") from error
        version = parse_libreoffice_version(completed.stdout + completed.stderr)
        if version is None:
            raise RecalculationError(
                "No se reconoce la versión de LibreOffice "
                f"({(completed.stdout or completed.stderr or 'sin salida').strip()})"
            )
        return version

    def check_minimum_version(self) -> tuple[int, ...]:
        version = self.installed_version()
        if version[:len(MINIMUM_LIBREOFFICE_VERSION)] < MINIMUM_LIBREOFFICE_VERSION:
            raise RecalculationError(
                f"LibreOffice {version_text(version)} es anterior a la versión mínima comprobada "
                f"({version_text(MINIMUM_LIBREOFFICE_VERSION)}). Actualízalo para generar el Excel."
            )
        return version

    def recalculate(self, workbook_path: Path, work_directory: Path) -> None:
        workbook_path = Path(workbook_path).resolve()
        if not workbook_path.is_file():
            raise RecalculationError(f"No existe el libro que se quiere recalcular: {workbook_path}")
        executable = self._find_executable()
        work_directory = Path(work_directory).resolve()
        output_directory = work_directory / "recalculated"
        output_directory.mkdir(parents=True, exist_ok=True)
        output_path = output_directory / workbook_path.name
        if output_path.exists():
            output_path.unlink()
        completed = None
        # LibreOffice crea rutas internas profundas. En expedientes ubicados
        # dentro de un worktree la ruta del perfil superaba MAX_PATH y soffice
        # terminaba con 0xC0000409 sin imprimir diagnóstico. El perfil es
        # efímero y puede vivir de forma segura en el directorio temporal corto.
        with tempfile.TemporaryDirectory(prefix="regularizacion-lo-") as profile_root:
            profile_directory = Path(profile_root).resolve()
            command = [
                str(executable),
                "--headless",
                f"-env:UserInstallation={profile_directory.as_uri()}",
                "--convert-to",
                "xlsx",
                "--outdir",
                str(output_directory),
                str(workbook_path),
            ]
            for attempt in range(2):
                try:
                    completed = subprocess.run(
                        command,
                        capture_output=True,
                        text=True,
                        timeout=self._timeout_seconds,
                        check=False,
                    )
                except subprocess.TimeoutExpired as error:
                    raise RecalculationError(
                        f"LibreOffice superó el límite de {self._timeout_seconds} segundos al recalcular"
                    ) from error
                except OSError as error:
                    raise RecalculationError(f"No se pudo iniciar LibreOffice: {error}") from error
                if completed.returncode != 0 or output_path.is_file():
                    break
                # Algunas instalaciones de Windows terminan la primera invocación
                # mientras aún inicializan el perfil aislado. Una segunda llamada
                # al mismo perfil completa la conversión de forma determinista.
                if attempt == 0:
                    time.sleep(0.2)
        assert completed is not None
        if completed.returncode != 0 or not output_path.is_file():
            details = (completed.stderr or completed.stdout or "sin diagnóstico").strip()
            raise RecalculationError(
                f"LibreOffice no pudo recalcular el libro ({details})"
            )
        os.replace(output_path, workbook_path)
