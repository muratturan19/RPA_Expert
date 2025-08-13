"""OCR engine using pytesseract for screen text detection."""

from __future__ import annotations

import time
import re
import unicodedata
from datetime import datetime
from typing import Optional, Tuple, Iterable
from pathlib import Path
from difflib import SequenceMatcher

import numpy as np
import pandas as pd
import pytesseract
import pyautogui
import pygetwindow as gw
import cv2
from PIL import Image, ImageDraw

import easyocr

from .config import (
    OCR_CONFIDENCE,
    OCR_LANGUAGE,
    OCR_TESSERACT_CONFIG,
    OCR_FUZZY_THRESHOLD,
)
from .logger import get_logger

logger = get_logger(__name__)


def demojibake(s: str) -> str:
    """Fix mojibake by re-decoding Latin-1 bytes as UTF-8."""
    try:
        return s.encode("latin-1").decode("utf-8")
    except Exception:
        return s


def normalize_tr(s: str) -> str:
    """Normalize Turkish text for case-insensitive comparisons.

    Converts Turkish-specific characters to their closest ASCII
    equivalents, collapses consecutive whitespace into a single space
    and lowercases the result. Other characters are preserved.
    """

    s = demojibake(s)
    s = unicodedata.normalize("NFKD", s)
    s = s.translate(str.maketrans("İIıŞşĞğÇçÖöÜü", "IIiSsGgCcOoUu"))
    s = re.sub(r"[-–—−-]", "-", s)
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s


def flexible_text_match(a: str, b: str, threshold: float = 0.8) -> bool:
    """Perform exact, partial and fuzzy matching between two strings.

    Tries both upper- and lower-case variants, treats Turkish ``İ``/``i``
    as equivalent and handles common OCR confusions between ``Ö`` and
    ``G``/``0`` characters.
    """

    def canonical_variants(text: str) -> set[str]:
        forms = {text, text.lower(), text.upper()}
        variants: set[str] = set()
        for form in forms:
            norm = normalize_tr(form)
            norm = norm.replace("0", "o").replace("g", "o")
            variants.add(norm)
        return variants

    a_vars = canonical_variants(a)
    b_vars = canonical_variants(b)

    for av in a_vars:
        for bv in b_vars:
            if av == bv or av in bv or bv in av:
                return True
            if SequenceMatcher(None, av, bv).ratio() >= threshold:
                return True
    return False


class OCREngine:
    def __init__(self, debug: bool = False):
        self.lang = OCR_LANGUAGE
        self.debug = debug
        self.tesseract_lang = "tur"
        self.easyocr_reader = easyocr.Reader(["tr", "en"], gpu=False)

        # Legacy attributes for backward compatibility
        self.use_easyocr = False
        self.reader = None

        # Create base debug directory and a timestamped run directory
        self.debug_root = Path("debug_screenshots")
        self.debug_root.mkdir(exist_ok=True)
        run_name = datetime.now().strftime("run_%Y%m%d_%H%M%S")
        self.run_dir = self.debug_root / run_name
        self.run_dir.mkdir(exist_ok=True)
        self.step = 0
        self.log_file = self.run_dir / "ocr_log.txt"

    def _save_debug_image(self, img, name: str) -> None:
        """Save screenshot for debugging purposes inside the run directory."""
        try:
            img.save(self.run_dir / f"{name}.png")
        except Exception as exc:
            logger.error("Failed to save debug image: %s", exc)

    def _screenshot(self, region=None, step_name: str = "step", region_pad: int = 0):
        try:
            self.step += 1
            step_label = f"step{self.step:02d}_{step_name}"

            windows = gw.getWindowsWithTitle("Preston")
            if windows:
                windows[0].activate()
                time.sleep(0.3)

            # Capture full screen for region overlay
            full_img = pyautogui.screenshot()
            if region:
                x, y, w, h = region
                if region_pad:
                    x = max(0, x - region_pad)
                    y = max(0, y - region_pad)
                    w += region_pad * 2
                    h += region_pad * 2
                raw_img = full_img.crop((x, y, x + w, y + h))
            else:
                raw_img = full_img

            processed_img = self._preprocess_image(raw_img)

            # Save raw and processed images
            raw_img.save(self.run_dir / f"{step_label}_raw.png")
            processed_img.save(self.run_dir / f"{step_label}_processed.png")

            if self.use_easyocr and self.reader:
                results = self.reader.readtext(np.array(processed_img))
                data = []
                text_lines = []
                for bbox, text, conf in results:
                    x_coords = [pt[0] for pt in bbox]
                    y_coords = [pt[1] for pt in bbox]
                    left = int(min(x_coords))
                    top = int(min(y_coords))
                    width = int(max(x_coords) - left)
                    height = int(max(y_coords) - top)
                    data.append(
                        {
                            "left": left,
                            "top": top,
                            "width": width,
                            "height": height,
                            "text": text,
                            "conf": conf * 100,
                        }
                    )
                    text_lines.append(text)
                df = pd.DataFrame(data)
                df.sort_values("top", inplace=True)
                line_num = 0
                last_top = -9999
                for idx, row in df.iterrows():
                    if row.top - last_top > 10:
                        line_num += 1
                        last_top = row.top
                    df.at[idx, "line_num"] = line_num
                ocr_text = "\n".join(text_lines)
            else:
                ocr_text = pytesseract.image_to_string(
                    processed_img, lang="tur+eng", config=OCR_TESSERACT_CONFIG
                )
                df = (
                    pytesseract.image_to_data(
                        processed_img,
                        lang="tur+eng",
                        config=OCR_TESSERACT_CONFIG,
                        output_type=pytesseract.Output.DATAFRAME,
                    ).dropna(subset=["text"])
                )

            with open(
                self.run_dir / f"{step_label}_ocr_result.txt", "w", encoding="utf-8"
            ) as f:
                f.write(ocr_text)
            df.to_csv(
                self.run_dir / f"{step_label}_ocr_data.csv",
                index=False,
                encoding="utf-8",
            )

            # Log texts, confidences and coordinates
            # Convert OCR engine coordinates (which are based on the scaled
            # ``processed_img``) back to the original screen space so that the
            # debug log reflects real cursor locations.
            scale_x = processed_img.width / raw_img.width if raw_img.width else 1
            scale_y = processed_img.height / raw_img.height if raw_img.height else 1
            with open(self.log_file, "a", encoding="utf-8") as log:
                for row in df.itertuples(index=False):
                    reg_left = row.left / scale_x
                    reg_top = row.top / scale_y
                    reg_w = row.width / scale_x
                    reg_h = row.height / scale_y
                    abs_left = int(reg_left + x) if region else int(reg_left)
                    abs_top = int(reg_top + y) if region else int(reg_top)
                    log.write(
                        f"{step_label}: {row.text} (conf={row.conf}, x={abs_left}, y={abs_top}, w={int(reg_w)}, h={int(reg_h)})\n"
                    )

            # Overlay region rectangle
            overlay = full_img.copy()
            if region:
                draw = ImageDraw.Draw(overlay)
                draw.rectangle(
                    [x, y, x + w, y + h], outline="red", width=2
                )
            overlay.save(self.run_dir / f"{step_label}_search_region.png")

            region_used = (x, y, w, h) if region else None
            return processed_img, df, step_label, region_used
        except Exception as exc:
            logger.error("Screenshot failed: %s", exc)
            return None, pd.DataFrame(), "", None

    @staticmethod
    def _preprocess_image(img: Image.Image) -> Image.Image:
        gray = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2GRAY)
        gray = cv2.resize(gray, None, fx=3, fy=3, interpolation=cv2.INTER_CUBIC)
        gray = cv2.GaussianBlur(gray, (5, 5), 0)
        _, thresh = cv2.threshold(
            gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU
        )
        kernel = np.ones((3, 3), np.uint8)
        opened = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel)
        closed = cv2.morphologyEx(opened, cv2.MORPH_CLOSE, kernel)
        return Image.fromarray(closed)

    @staticmethod
    def _normalize(text: str) -> str:
        """Thin wrapper around normalize_tr for backward compatibility."""
        return normalize_tr(text)

    def find_word_pair(
        self,
        window_rect: Tuple[int, int, int, int],
        left_word: str = "finans",
        right_word: str = "izle",
        max_gap: int = 300,
        conf_min: float = 30,
        region_pad: int = 0,
    ) -> Optional[Tuple[int, int, int, int]]:
        """Locate two words on the same line within a given pixel gap."""
        img, df, _, used_region = self._screenshot(
            region=window_rect,
            step_name=f"word_pair_{left_word}_{right_word}",
            region_pad=region_pad,
        )
        if img is None or df.empty:
            return None
        df["conf"] = df["conf"].astype(float)
        df["ntext"] = df["text"].map(self._normalize)
        df = df[df["conf"] >= conf_min]

        left_norm = self._normalize(left_word)
        right_norm = self._normalize(right_word)
        left_variants = (
            ["Finans", "finans", "FINANS"]
            if left_norm == "finans"
            else [left_word, left_word.lower(), left_word.upper()]
        )
        right_variants = (
            ["İzle", "izle", "IZLE", "Izle"]
            if right_norm == "izle"
            else [right_word, right_word.lower(), right_word.upper()]
        )
        left_targets = {self._normalize(w) for w in left_variants}
        right_targets = {self._normalize(w) for w in right_variants}

        # Search left word first, then look for the right word on the same line
        left_tokens = df[df.ntext.isin(left_targets)].sort_values(["line_num", "left"])
        for L in left_tokens.itertuples():
            line_tokens = df[(df.line_num == L.line_num) & (df.left > L.left)]
            line_tokens = line_tokens[line_tokens.left - L.left < max_gap]
            right_candidates = line_tokens[line_tokens.ntext.isin(right_targets)]
            if not right_candidates.empty:
                R = right_candidates.iloc[0]

                # Scale factors between OCR image size and original region
                region_w = used_region[2] if used_region else window_rect[2]
                region_h = used_region[3] if used_region else window_rect[3]
                scale_x = img.width / region_w if region_w else 1
                scale_y = img.height / region_h if region_h else 1

                # Coordinates relative to the search region
                rel_x = min(L.left, R.left) / scale_x
                rel_y = min(L.top, R.top) / scale_y
                rel_w = (
                    max(L.left + L.width, R.left + R.width) - min(L.left, R.left)
                ) / scale_x
                rel_h = (
                    max(L.top + L.height, R.top + R.height) - min(L.top, R.top)
                ) / scale_y

                # Absolute coordinates on the screen
                base_x = used_region[0] if used_region else window_rect[0]
                base_y = used_region[1] if used_region else window_rect[1]
                abs_x = int(rel_x + base_x)
                abs_y = int(rel_y + base_y)
                abs_w = int(rel_w)
                abs_h = int(rel_h)

                logger.debug(
                    "Word pair '%s' '%s' relative coords=%s absolute coords=%s",
                    left_word,
                    right_word,
                    (int(rel_x), int(rel_y), int(rel_w), int(rel_h)),
                    (abs_x, abs_y, abs_w, abs_h),
                )
                try:
                    with open(self.log_file, "a", encoding="utf-8") as log:
                        log.write(
                            f"word_pair {left_word} {right_word}: rel=({int(rel_x)}, {int(rel_y)}, {int(rel_w)}, {int(rel_h)}) abs=({abs_x}, {abs_y}, {abs_w}, {abs_h})\n"
                        )
                except Exception:
                    pass
                return abs_x, abs_y, abs_w, abs_h
        if self.debug:
            self._save_debug_image(img, f"pair_not_found_{left_word}_{right_word}")
        return None

    def click_word_pair(
        self,
        window_rect: Tuple[int, int, int, int],
        left_word: str = "finans",
        right_word: str = "izle",
        max_gap: int = 300,
        conf_min: float = 30,
        region_pad: int = 0,
    ) -> bool:
        """Find a word pair and click the centre of their combined bounding box."""
        bbox = self.find_word_pair(
            window_rect,
            left_word=left_word,
            right_word=right_word,
            max_gap=max_gap,
            conf_min=conf_min,
            region_pad=region_pad,
        )
        if bbox:
            x, y, w, h = bbox
            pyautogui.click(x + w // 2, y + h // 2)
            time.sleep(0.1)
            return True
        logger.error("Word pair '%s' and '%s' not found on screen", left_word, right_word)
        return False

    def _find_text_engine(
        self,
        targets: list[str],
        variants: list[str],
        region,
        confidence: float,
        normalize: bool,
        region_pad: int,
        use_easyocr: bool,
        found_texts: Optional[list[str]] = None,
    ) -> Optional[Tuple[int, int, int, int]]:
        orig_use_easyocr = self.use_easyocr
        orig_reader = self.reader
        self.use_easyocr = use_easyocr
        self.reader = self.easyocr_reader if use_easyocr else None
        try:
            img, df, _, used_region = self._screenshot(
                region=region, step_name="find_text", region_pad=region_pad
            )
        finally:
            self.use_easyocr = orig_use_easyocr
            self.reader = orig_reader
        if img is None or df.empty:
            return None
        df["conf"] = df["conf"].astype(float)
        df = df[df["conf"] >= confidence * 100]
        lines: dict = {}
        for row in df.itertuples(index=False):
            if not str(row.text).strip():
                continue
            page_num = getattr(row, "page_num", 0)
            block_num = getattr(row, "block_num", 0)
            par_num = getattr(row, "par_num", 0)
            line_num = getattr(row, "line_num", 0)
            key = (page_num, block_num, par_num, line_num)
            line = lines.setdefault(
                key,
                {"words": [], "left": [], "top": [], "right": [], "bottom": []},
            )
            line["words"].append(row.text)
            left, top, width, height = row.left, row.top, row.width, row.height
            line["left"].append(left)
            line["top"].append(top)
            line["right"].append(left + width)
            line["bottom"].append(top + height)
        for line in lines.values():
            line_text = " ".join(line["words"])
            if found_texts is not None:
                found_texts.append(line_text)
            line_norm = (
                self._normalize(line_text) if normalize else line_text.casefold()
            )
            for target in targets:
                ratio = SequenceMatcher(None, line_norm, target).ratio()
                if target in line_norm or ratio >= OCR_FUZZY_THRESHOLD:
                    x = min(line["left"])
                    y = min(line["top"])
                    w = max(line["right"]) - x
                    h = max(line["bottom"]) - y
                    if used_region:
                        x += used_region[0]
                        y += used_region[1]
                    return x, y, w, h
        if self.debug and variants:
            miss = self._normalize(variants[0]) if normalize else variants[0].casefold()
            self._save_debug_image(img, f"not_found_{miss}")
        return None

    def find_text_on_screen(
        self,
        text: Iterable[str] | str,
        region=None,
        confidence: float = OCR_CONFIDENCE,
        normalize: bool = True,
        region_pad: int = 0,
        texts_out: Optional[list[str]] = None,
    ) -> Optional[Tuple[int, int, int, int]]:
        """Find text coordinates using EasyOCR first, then fall back to Tesseract."""

        variants = [text] if isinstance(text, str) else list(text)
        targets = [self._normalize(v) if normalize else v.casefold() for v in variants]

        for use_easyocr, name in ((True, "easyocr"), (False, "tesseract")):
            try:
                bbox = self._find_text_engine(
                    targets,
                    variants,
                    region,
                    confidence,
                    normalize,
                    region_pad,
                    use_easyocr,
                    texts_out,
                )
            except Exception as exc:
                logger.exception("%s engine failed: %s", name, exc)
                bbox = None
            if bbox:
                return bbox
        return None

    def click_text(
        self,
        text: Iterable[str] | str,
        offset_x: int = 0,
        offset_y: int = 0,
        region=None,
        region_pad: int = 0,
    ) -> bool:
        """Click on found text or any of its variants."""
        variants = [text] if isinstance(text, str) else list(text)
        for variant in variants:
            coords = self.find_text_on_screen(
                variant, region=region, region_pad=region_pad
            )
            if coords:
                x, y, w, h = coords
                pyautogui.click(x + w // 2 + offset_x, y + h // 2 + offset_y)
                time.sleep(0.1)
                return True
        logger.error("Text '%s' not found on screen", text)
        if self.debug:
            logger.debug("Saved debug screenshot for '%s'", text)
        return False

    def wait_for_text(
        self,
        text: Iterable[str] | str,
        timeout: float = 10,
        region=None,
        region_pad: int = 0,
        confidence: float = OCR_CONFIDENCE,
    ) -> bool:
        """Wait until text appears on screen."""
        end_time = time.time() + timeout
        while time.time() < end_time:
            if self.find_text_on_screen(
                text, region=region, region_pad=region_pad, confidence=confidence
            ):
                return True
            time.sleep(0.5)
        logger.error("Timeout waiting for text: %s", text)
        return False

    def find_word_pair_tesseract(self, img, left_word: str, right_word: str, debug_dir: Path):
        df = (
            pytesseract.image_to_data(
                img, lang=self.tesseract_lang, output_type=pytesseract.Output.DATAFRAME
            ).dropna(subset=["text"])
        )
        debug_dir.mkdir(exist_ok=True)
        annotated = img.copy()
        draw = ImageDraw.Draw(annotated)
        log_path = debug_dir / "log.txt"
        with open(log_path, "w", encoding="utf-8") as log:
            for row in df.itertuples(index=False):
                draw.rectangle(
                    [row.left, row.top, row.left + row.width, row.top + row.height],
                    outline="red",
                    width=1,
                )
                log.write(
                    f"{row.text}\t{row.conf}\t{row.left},{row.top},{row.width},{row.height}\n"
                )
        annotated.save(debug_dir / "annotated.png")
        df["ntext"] = df["text"].map(self._normalize)
        df.sort_values("top", inplace=True)
        line_num = 0
        last_top = -9999
        for idx, row in df.iterrows():
            if row.top - last_top > 10:
                line_num += 1
                last_top = row.top
            df.at[idx, "line_num"] = line_num
        left_variants = ["Finans", "finans", "FINANS"]
        right_variants = ["İzle", "izle", "IZLE", "Izle"]
        left_targets = {self._normalize(w) for w in left_variants}
        right_targets = {self._normalize(w) for w in right_variants}

        left_tokens = df[df.ntext.isin(left_targets)].sort_values(["line_num", "left"])
        for _, L in left_tokens.iterrows():
            line_tokens = df[(df.line_num == L.line_num) & (df.left > L.left)]
            right_candidates = line_tokens[line_tokens.ntext.isin(right_targets)]
            if not right_candidates.empty:
                R = right_candidates.iloc[0]
                x = min(L.left, R.left)
                y = min(L.top, R.top)
                w = max(L.left + L.width, R.left + R.width) - x
                h = max(L.top + L.height, R.top + R.height) - y
                return int(x), int(y), int(w), int(h)
        return None

    def find_word_pair_easyocr(self, img, left_word: str, right_word: str, debug_dir: Path):
        try:
            results = self.easyocr_reader.readtext(np.array(img))
        except Exception as e:
            logger.error(f"EasyOCR failed: {e}")
            results = []
        logger.info(f"EasyOCR found {len(results)} items")
        data = []
        debug_dir.mkdir(exist_ok=True)
        annotated = img.copy()
        draw = ImageDraw.Draw(annotated)
        log_path = debug_dir / "log.txt"
        with open(log_path, "w", encoding="utf-8") as log:
            for bbox, text, conf in results:
                x_coords = [pt[0] for pt in bbox]
                y_coords = [pt[1] for pt in bbox]
                left = int(min(x_coords))
                top = int(min(y_coords))
                width = int(max(x_coords) - left)
                height = int(max(y_coords) - top)
                draw.rectangle(
                    [left, top, left + width, top + height], outline="red", width=1
                )
                log.write(f"{text}\t{conf}\t{left},{top},{width},{height}\n")
                data.append(
                    {
                        "left": left,
                        "top": top,
                        "width": width,
                        "height": height,
                        "text": text,
                        "conf": conf * 100,
                    }
                )
        annotated.save(debug_dir / "annotated.png")
        df = pd.DataFrame(data)
        if df.empty:
            return None
        df["ntext"] = df["text"].map(self._normalize)
        df.sort_values("top", inplace=True)
        line_num = 0
        last_top = -9999
        for idx, row in df.iterrows():
            if row.top - last_top > 10:
                line_num += 1
                last_top = row.top
            df.at[idx, "line_num"] = line_num
        left_variants = ["Finans", "finans", "FINANS"]
        right_variants = ["İzle", "izle", "IZLE", "Izle"]
        left_targets = {self._normalize(w) for w in left_variants}
        right_targets = {self._normalize(w) for w in right_variants}

        left_tokens = df[df.ntext.isin(left_targets)].sort_values(["line_num", "left"])
        for _, L in left_tokens.iterrows():
            line_tokens = df[(df.line_num == L.line_num) & (df.left > L.left)]
            right_candidates = line_tokens[line_tokens.ntext.isin(right_targets)]
            if not right_candidates.empty:
                R = right_candidates.iloc[0]
                x = min(L.left, R.left)
                y = min(L.top, R.top)
                w = max(L.left + L.width, R.left + R.width) - x
                h = max(L.top + L.height, R.top + R.height) - y
                return int(x), int(y), int(w), int(h)
        return None

