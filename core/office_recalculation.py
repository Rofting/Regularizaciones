"""Recálculo controlado de libros mediante LibreOffice o una doble de pruebas."""

from __future__ import annotations

import os
import shutil
import subprocess
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
        if self._executable is not None:
            if self._executable.is_file():
                return self._executable
            raise RecalculationError(
                "LibreOffice no está disponible en la ruta configurada. "
                "Instálalo o selecciona su ejecutable soffice."
            )
        for command in ("soffice", "libreoffice"):
            located = shutil.which(command)
            if located:
                return Path(located)
        candidates: list[Path] = []
        for variable in ("ProgramFiles", "ProgramFiles(x86)"):
            base = os.environ.get(variable)
            if base:
                candidates.append(Path(base) / "LibreOffice" / "program" / "soffice.exe")
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
        profile_directory = work_directory / "libreoffice_profile"
        output_directory.mkdir(parents=True, exist_ok=True)
        profile_directory.mkdir(parents=True, exist_ok=True)
        output_path = output_directory / workbook_path.name
        if output_path.exists():
            output_path.unlink()
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
        if completed.returncode != 0 or not output_path.is_file():
            details = (completed.stderr or completed.stdout or "sin diagnóstico").strip()
            raise RecalculationError(
                f"LibreOffice no pudo recalcular el libro ({details})"
            )
        os.replace(output_path, workbook_path)
