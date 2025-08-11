"""Streamlit user interface for Preston RPA."""

from __future__ import annotations

import io
import threading
from pathlib import Path
from typing import List, Dict

import streamlit as st

from .excel_processor import process_excel_file
from .preston_automation import PrestonRPA
from .logger import get_logger

logger = get_logger(__name__)


def run_automation(data: List[Dict[str, object]], simulator_path: str, progress_placeholder):
    rpa = PrestonRPA()
    total = len(data)
    for idx, entry in enumerate(data, start=1):
        if not rpa.running:
            break
        rpa.execute_workflow(entry)
        progress_placeholder.progress(idx / total)
    rpa.stop()


def main():
    st.set_page_config(page_title="Preston RPA", layout="wide")
    st.title("Preston RPA Automation")

    with st.sidebar:
        st.header("Configuration")
        uploaded_file = st.file_uploader("Excel file", type=["xls", "xlsx"])
        simulator_path = st.text_input("Simulator Path", value=str(Path(__file__).parent / "Preston.html"))
        start_button = st.button("Start RPA", type="primary")

    log_box = st.empty()
    progress_placeholder = st.progress(0.0)

    if start_button and uploaded_file is not None:
        with io.BytesIO(uploaded_file.read()) as buffer:
            excel_path = Path("uploaded.xlsx")
            excel_path.write_bytes(buffer.getvalue())
            data = process_excel_file(excel_path)
        threading.Thread(target=run_automation, args=(data, simulator_path, progress_placeholder), daemon=True).start()
        st.success("Automation started")

    log_box.text(Path(__file__).with_name("automation.log").read_text() if Path(__file__).with_name("automation.log").exists() else "")


if __name__ == "__main__":
    main()
