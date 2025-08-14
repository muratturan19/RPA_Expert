"""Streamlit user interface for Preston RPA."""

from __future__ import annotations

import io
import threading
import time
from queue import Queue
from pathlib import Path
from typing import List, Dict

import streamlit as st

if __package__ in (None, ""):
    import sys
    sys.path.append(str(Path(__file__).resolve().parent.parent))
    from preston_rpa.excel_processor import process_excel_file
    from preston_rpa.preston_automation import PrestonRPA, focus_preston_window
    from preston_rpa.logger import get_logger
else:
    from .excel_processor import process_excel_file
    from .preston_automation import PrestonRPA, focus_preston_window
    from .logger import get_logger

logger = get_logger(__name__)


def run_automation(data: List[Dict[str, object]], simulator_path: str, progress_queue: Queue):
    rpa = PrestonRPA()
    total = len(data)
    try:
        for idx, entry in enumerate(data, start=1):
            if not rpa.running:
                break
            rpa.execute_workflow(entry)
            progress_queue.put(idx / total)
    except Exception as exc:
        logger.exception("Automation failed: %s", exc)
        progress_queue.put(("error", str(exc)))
    finally:
        rpa.stop()
        progress_queue.put(1.0)


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
        focus_preston_window(simulator_path)
        progress_queue: Queue = Queue()
        thread = threading.Thread(
            target=run_automation,
            args=(data, simulator_path, progress_queue),
            daemon=True,
        )
        thread.start()
        error_message = None
        while thread.is_alive() or not progress_queue.empty():
            try:
                progress = progress_queue.get(timeout=0.1)
                if isinstance(progress, tuple) and progress[0] == "error":
                    error_message = progress[1]
                    st.error(error_message)
                    break
                else:
                    progress_placeholder.progress(progress)
            except Exception:
                pass
            time.sleep(0.1)
        if error_message is None:
            st.success("Automation finished")

    log_path = Path(__file__).with_name("automation.log")
    if log_path.exists():
        try:
            log_content = log_path.read_text(encoding="utf-8")
        except Exception as exc:
            st.error(f"Failed to read log file: {exc}")
            log_content = ""
    else:
        log_content = ""
    log_box.text(log_content)


if __name__ == "__main__":
    main()
