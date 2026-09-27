"""Recálculo controlado de libros mediante LibreOffice o una doble de pruebas."""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
import time
from pathlib import Path
from typing import Callable, Protocol


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
