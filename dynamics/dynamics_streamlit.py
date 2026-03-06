# dynamics_streamlit.py
from __future__ import annotations

import os
import tempfile
import datetime as dt
from typing import Optional, Tuple

from streamlit_app.dynamics import dynamics


def build_dynamics_report_streamlit(
    db_excel_bytes: bytes,
    report_kind: str,  # "Ежедневный" | "Недельный" | "Месячный"
    gz_year: int,
    plan_count: int,
    plan_area: float,
    split_by_ogx: bool = False,
    split_by_itp: bool = False,
) -> Tuple[bytes, str]:
    """
    bytes Excel -> временный input.xlsx -> генерим отчёт -> читаем bytes -> отдаём в Streamlit
    """

    if not db_excel_bytes:
        raise ValueError("Пустой Excel (db_excel_bytes).")

    # чтобы не пытался открывать файл на сервере/ПК
    if hasattr(dynamics, "open_file_crossplatform"):
        dynamics.open_file_crossplatform = lambda _: None

    with tempfile.TemporaryDirectory() as tmpdir:
        in_path = os.path.join(tmpdir, "input.xlsx")
        with open(in_path, "wb") as f:
            f.write(db_excel_bytes)

        # генератор может писать файл в текущую папку, поэтому делаем cwd=tmpdir
        old_cwd = os.getcwd()
        os.chdir(tmpdir)

        out_path: Optional[str] = None
        try:
            gen = dynamics.ReportGenerator(in_path)

            if report_kind == "Ежедневный":
                out_path = gen.generate_daily(
                    gz_year, plan_count, plan_area,
                    split_by_ogx=split_by_ogx, split_by_itp=split_by_itp
                )
            elif report_kind == "Недельный":
                out_path = gen.generate_weekly_combined(
                    gz_year, plan_count, plan_area,
                    split_by_ogx=split_by_ogx, split_by_itp=split_by_itp
                )
            elif report_kind == "Месячный":
                out_path = gen.generate_monthly_combined(
                    gz_year, plan_count, plan_area,
                    split_by_ogx=split_by_ogx, split_by_itp=split_by_itp
                )
            else:
                raise ValueError(f"Неизвестный тип отчёта: {report_kind}")

            if not out_path or not os.path.exists(out_path):
                # запасной поиск: любой xlsx кроме input.xlsx
                candidates = [p for p in os.listdir(tmpdir) if p.endswith(".xlsx") and p != "input.xlsx"]
                if not candidates:
                    raise FileNotFoundError("Не найден созданный Excel-файл отчёта в tmp-папке.")
                out_path = os.path.join(tmpdir, max(candidates, key=lambda p: os.path.getmtime(os.path.join(tmpdir, p))))

            with open(out_path, "rb") as f:
                excel_bytes = f.read()

            filename = f"Динамика_{report_kind}_ГЗ-{gz_year}_{dt.datetime.now().strftime('%d.%m.%Y_%H-%M')}.xlsx"
            return excel_bytes, filename

        finally:
            os.chdir(old_cwd)