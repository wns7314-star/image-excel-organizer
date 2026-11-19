#!/usr/bin/env python3
"""이미지 파일을 파일명 규칙에 따라 Excel 표로 정리하는 GUI 프로그램.

핵심 원칙
---------
* 확장자로 거르지 않고 선택된 모든 파일의 실제 내용을 이미지 디코더로 확인한다.
* 파일명 첫 구간을 행 그룹으로, 나머지 숫자들을 행 내부 정렬 키로 사용한다.
* 원본 픽셀을 축소 저장하지 않고, Excel 안에서 표시 크기만 셀에 맞춘다.
* 각 이미지를 고유한 임시 PNG로 만든 뒤 Workbook이 닫힐 때까지 보존한다.
* 완성된 XLSX의 ZIP/XML/이미지 앵커 수를 검증한 뒤에만 최종 파일로 옮긴다.

필수 패키지: Pillow, XlsxWriter
선택 패키지: pillow-heif(HEIC/HEIF), CairoSVG(SVG/SVGZ), rawpy(카메라 RAW)
"""

from __future__ import annotations

import gzip
import json
import math
import os
import queue
import re
import tempfile
import threading
import unicodedata
import uuid
import zipfile
from collections import OrderedDict
from dataclasses import dataclass
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Callable, Iterable, Sequence
from xml.etree import ElementTree

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except ImportError as exc:  # pragma: no cover - 공식 Windows Python에는 기본 포함
    raise SystemExit("tkinter가 없습니다. python.org의 Windows Python을 설치해 주세요.") from exc

try:
    import xlsxwriter
except ImportError as exc:
    raise SystemExit(
        "XlsxWriter가 없습니다. 터미널에서 'python -m pip install XlsxWriter'를 실행해 주세요."
    ) from exc

try:
    from PIL import Image, ImageOps, UnidentifiedImageError
except ImportError as exc:
    raise SystemExit(
        "Pillow가 없습니다. 터미널에서 'python -m pip install Pillow'를 실행해 주세요."
    ) from exc

try:  # 선택 기능: 탐색기에서 파일 드래그앤드롭
    from tkinterdnd2 import DND_FILES, TkinterDnD

    DND_AVAILABLE = True
except (ImportError, OSError):  # pragma: no cover - 선택 패키지
    DND_FILES = None
    TkinterDnD = None
    DND_AVAILABLE = False


APP_TITLE = "이미지 → Excel 자동 정리"
EXCEL_FONT_NAME = "맑은 고딕"
IMAGE_COLUMN_WIDTH = 17.0
IMAGE_ROW_HEIGHT = 80.0  # Excel 행 높이 단위: point
VALIDATION_GROUP_WIDTH = 10.0
VALIDATION_FILENAME_WIDTH = 15.0
VALIDATION_GAP_WIDTH = 2.5
TITLE_ROW_HEIGHT = 26.0
HEADER_ROW_HEIGHT = 24.0
DATA_START_ROW = 2  # XlsxWriter의 0 기반 행: Excel에서는 3행
MAX_EXCEL_ROWS = 1_048_576
MAX_EXCEL_COLUMNS = 16_384
SETTINGS_FILENAME = "settings.json"
SIZE_PRESETS: dict[str, tuple[float, float] | None] = {
    "기본 (17 × 80)": (17.0, 80.0),
    "작게 (12 × 60)": (12.0, 60.0),
    "중간 (20 × 100)": (20.0, 100.0),
    "크게 (25 × 130)": (25.0, 130.0),
    "사용자 설정": None,
}

NATURAL_PART_RE = re.compile(r"(\d+)")
LEADING_GROUP_RE = re.compile(r"^([^\d]*\d+)")
INVALID_WINDOWS_FILENAME_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
WINDOWS_RESERVED_BASENAMES = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    *(f"COM{number}" for number in range(1, 10)),
    *(f"LPT{number}" for number in range(1, 10)),
}

ProgressCallback = Callable[[int, int, str], None]
CancelCheck = Callable[[], bool]


@dataclass(frozen=True)
class ImageItem:
    path: Path
    group: str
    sequence_text: str


@dataclass(frozen=True)
class ImageResult:
    order: int
    group: str
    position_in_group: int
    path: Path
    success: bool
    pixel_width: int | None = None
    pixel_height: int | None = None
    frame_count: int | None = None
    source_format: str | None = None
    note: str = ""


@dataclass(frozen=True)
class GenerationSummary:
    output_path: Path
    selected_count: int
    inserted_count: int
    failed_count: int
    group_count: int
    results: tuple[ImageResult, ...]


class ImageConversionError(RuntimeError):
    """선택 파일을 Excel용 PNG로 변환할 수 없을 때 사용한다."""


class GenerationCancelled(RuntimeError):
    """사용자가 진행 중인 Excel 생성을 취소했을 때 사용한다."""


def register_optional_decoders() -> dict[str, bool]:
    """설치되어 있는 선택 이미지 디코더를 등록하고 상태를 반환한다."""

    status = {"HEIC/HEIF": False, "SVG/SVGZ": False, "카메라 RAW": False}

    try:
        import pillow_heif

        pillow_heif.register_heif_opener()
        status["HEIC/HEIF"] = True
    except (ImportError, OSError):
        pass

    try:
        import cairosvg  # noqa: F401

        status["SVG/SVGZ"] = True
    except (ImportError, OSError):
        pass

    try:
        import rawpy  # noqa: F401

        status["카메라 RAW"] = True
    except (ImportError, OSError):
        pass

    return status


OPTIONAL_DECODERS = register_optional_decoders()


def normalize_filename_text(text: str) -> str:
    """전각 문자 등을 통일하고 앞뒤 공백을 제거한다."""

    return unicodedata.normalize("NFKC", text).strip()


def natural_sort_key(text: str) -> tuple[tuple[int, object], ...]:
    """숫자 덩어리를 정수로 비교하는, 서로 타입 충돌이 없는 정렬 키를 만든다.

    예: ``1_1_2``는 ``1_1_10``보다 먼저 온다. 숫자가 같으면 앞의 0 개수와
    대소문자 차이가 결과를 흔들지 않도록 정규화된 원문으로 한 번 더 비교한다.
    """

    normalized = normalize_filename_text(text).casefold()
    key: list[tuple[int, object]] = []
    for part in NATURAL_PART_RE.split(normalized):
        if not part:
            continue
        if part.isdigit():
            key.append((0, int(part)))
        else:
            key.append((1, part))
    key.append((2, normalized))
    return tuple(key)


def sequence_sort_key(
    text: str,
) -> tuple[int, tuple[int, ...], tuple[tuple[int, object], ...]]:
    """행 내부 정렬 키를 만든다.

    같은 행에서는 구분자(``-``, ``_``, 공백)나 설명 문자가 숫자 순서를
    방해하지 않도록, 나머지 파일명에서 발견한 숫자들을 왼쪽부터 비교한다.
    숫자 배열이 완전히 같을 때만 정규화된 원문을 최종 동률 해소에 사용한다.

    예: ``1-10`` → ``(1, 10)``, ``1_2`` → ``(1, 2)``이므로 두 번째가 먼저다.
    """

    normalized = normalize_filename_text(text).casefold()
    numbers = tuple(int(value) for value in re.findall(r"\d+", normalized))
    if numbers:
        return 0, numbers, natural_sort_key(normalized)
    return 1, (), natural_sort_key(normalized)


def split_group_and_sequence(filename_or_stem: str) -> tuple[str, str]:
    """파일명을 ``(행 그룹, 행 안에서 정렬할 나머지 부분)``으로 분리한다.

    우선순위:
    1. 첫 영문자/한글/숫자가 아닌 문자 앞의 구간:
       ``1-01-2``, ``1.01.2``, ``1#01#2``는 모두 그룹 ``1``
    2. 그런 경계 문자가 없으면 선두의 문자+숫자:
       ``No1사진2`` → 그룹 ``No1``, 정렬 부분 ``사진2``
    3. 위 규칙이 없으면 전체 파일명을 그룹으로 사용하고 정렬 부분은 비운다.
    앞부분은 행을 정할 때만 사용하고, 같은 행 안의 순서는 나머지 부분의
    숫자들을 왼쪽부터 자연 정렬해 결정한다.
    """

    stem = normalize_filename_text(Path(filename_or_stem).stem)
    if not stem:
        return "(빈 파일명)", ""

    # 특정 구분자 목록에 의존하지 않는다. 문자/숫자가 아닌 첫 문자를
    # 그룹 경계로 보므로 -, _, 공백뿐 아니라 ., #, +, 괄호 등도 처리한다.
    for index, character in enumerate(stem):
        if index > 0 and not character.isalnum():
            sequence = stem[index:]
            # 연속된 경계 기호는 행 내부 숫자 정렬에 필요하지 않다.
            boundary_length = 0
            while boundary_length < len(sequence) and not sequence[boundary_length].isalnum():
                boundary_length += 1
            sequence = sequence[boundary_length:]
            return stem[:index], sequence

    match = LEADING_GROUP_RE.match(stem)
    if match:
        return match.group(1), stem[match.end() :]
    return stem, ""


def extract_group(filename_or_stem: str) -> str:
    """파일명에서 같은 행에 배치할 앞부분 그룹만 반환한다."""

    return split_group_and_sequence(filename_or_stem)[0]


def organize_paths(
    paths: Iterable[str | os.PathLike[str]],
    manual_order: Sequence[str | os.PathLike[str]] | None = None,
) -> list[tuple[str, list[ImageItem]]]:
    """중복 경로를 제거하고 그룹/파일명 자연 정렬 결과를 반환한다.

    ``manual_order``가 주어지면 행 그룹의 자연 정렬은 유지하면서 같은 행 그룹
    내부에서만 사용자가 지정한 순서를 우선한다. GUI 미리보기에서 위/아래로
    조정한 순서를 실제 Excel 출력에도 그대로 반영하기 위한 옵션이다.
    """

    unique: dict[str, Path] = {}
    for value in paths:
        path = Path(value).expanduser().resolve()
        identity = os.path.normcase(str(path))
        unique.setdefault(identity, path)

    items: list[ImageItem] = []
    for path in unique.values():
        group, sequence_text = split_group_and_sequence(path.name)
        items.append(ImageItem(path=path, group=group, sequence_text=sequence_text))

    items.sort(
        key=lambda item: (
            natural_sort_key(item.group),
            sequence_sort_key(item.sequence_text),
            natural_sort_key(item.path.stem),
            natural_sort_key(item.path.suffix),
            os.path.normcase(str(item.path)),
        )
    )

    grouped: "OrderedDict[str, list[ImageItem]]" = OrderedDict()
    group_identity_to_label: dict[str, str] = {}
    for item in items:
        group_identity = normalize_filename_text(item.group).casefold()
        label = group_identity_to_label.setdefault(group_identity, item.group)
        grouped.setdefault(label, []).append(item)

    if manual_order:
        rank: dict[str, int] = {}
        for index, value in enumerate(manual_order):
            path = Path(value).expanduser().resolve()
            rank.setdefault(os.path.normcase(str(path)), index)

        for group_items in grouped.values():
            natural_rank = {
                os.path.normcase(str(item.path.resolve())): index
                for index, item in enumerate(group_items)
            }
            group_items.sort(
                key=lambda item: (
                    rank.get(
                        os.path.normcase(str(item.path.resolve())),
                        len(rank) + natural_rank[os.path.normcase(str(item.path.resolve()))],
                    ),
                    natural_rank[os.path.normcase(str(item.path.resolve()))],
                )
            )

    return list(grouped.items())


def settings_path() -> Path:
    """사용자별 GUI 설정 파일 경로를 반환한다."""

    if os.name == "nt" and os.environ.get("APPDATA"):
        root = Path(os.environ["APPDATA"]) / "ImageExcelOrganizer"
    else:
        root = Path.home() / ".config" / "image_excel_organizer"
    return root / SETTINGS_FILENAME


def load_settings(path: Path | None = None) -> dict[str, object]:
    """설정 JSON을 안전하게 읽는다. 손상/부재 시 빈 설정을 반환한다."""

    target = path or settings_path()
    try:
        if not target.is_file():
            return {}
        data = json.loads(target.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, json.JSONDecodeError, UnicodeError):
        return {}


def save_settings(settings: dict[str, object], path: Path | None = None) -> None:
    """설정 JSON을 원자적으로 저장한다."""

    target = path or settings_path()
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = target.with_suffix(target.suffix + ".tmp")
    temporary.write_text(
        json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    os.replace(temporary, target)


def analyze_selection(
    paths: Sequence[str | os.PathLike[str]], *, check_missing: bool = False
) -> list[str]:
    """내보내기 전에 중복/누락 가능성을 보수적으로 검사한다.

    누락 검사는 선택 옵션이며, 같은 행 그룹 안에서 모든 파일의 숫자 토큰 개수가
    동일할 때 마지막 숫자만 연속 번호로 간주한다. 확정 오류가 아닌 '확인 권고'다.
    """

    warnings: list[str] = []
    basename_map: dict[str, list[Path]] = {}
    for raw in paths:
        path = Path(raw).expanduser().resolve()
        basename_map.setdefault(path.name.casefold(), []).append(path)
        if not path.is_file():
            warnings.append(f"파일을 찾을 수 없음: {path}")

    for same_name in basename_map.values():
        if len(same_name) > 1:
            locations = " / ".join(str(path.parent) for path in same_name)
            warnings.append(f"동일 파일명 {same_name[0].name}: {locations}")

    if not check_missing:
        return warnings

    for group_name, items in organize_paths(paths):
        token_lists = [re.findall(r"\d+", item.sequence_text) for item in items]
        if len(items) < 3 or not token_lists or any(not tokens for tokens in token_lists):
            continue
        token_count = len(token_lists[0])
        if token_count == 0 or any(len(tokens) != token_count for tokens in token_lists):
            continue

        # 앞 숫자 토큰이 동일한 하위 묶음별로 마지막 숫자의 누락만 찾는다.
        buckets: dict[tuple[int, ...], set[int]] = {}
        for tokens in token_lists:
            numeric = tuple(int(token) for token in tokens)
            buckets.setdefault(numeric[:-1], set()).add(numeric[-1])
        for prefix, values in buckets.items():
            if len(values) < 2:
                continue
            low, high = min(values), max(values)
            if high - low > 5000:  # 비정상적으로 큰 범위는 추정하지 않는다.
                continue
            missing = [value for value in range(low, high + 1) if value not in values]
            if missing:
                prefix_text = "-".join(str(value) for value in prefix) if prefix else ""
                missing_text = ", ".join(str(value) for value in missing[:20])
                suffix = " ..." if len(missing) > 20 else ""
                context = f" ({prefix_text})" if prefix_text else ""
                warnings.append(
                    f"행 그룹 {group_name}{context}: 번호 누락 가능성 {missing_text}{suffix}"
                )
    return warnings


def excel_column_width_to_pixels(width: float) -> int:
    """Excel 기본 글꼴(Calibri 11) 기준 열 너비를 픽셀로 변환한다.

    XlsxWriter가 개체 위치 계산에 사용하는 Excel 호환 변환식과 같다.
    너비 17은 124px이다.
    """

    max_digit_width = 7
    padding = 5
    if width < 1:
        return int(width * (max_digit_width + padding) + 0.5)
    return int(width * max_digit_width + 0.5) + padding


def excel_row_height_to_pixels(height: float) -> int:
    """Excel 행 높이(point)를 XlsxWriter 개체 배치용 픽셀로 변환한다."""

    return int(4.0 / 3.0 * height)


def validate_image_cell_size(
    column_width: float | str,
    row_height: float | str,
) -> tuple[float, float]:
    """GUI/함수 입력으로 받은 이미지 셀 크기를 검증하고 숫자로 반환한다.

    Excel/XlsxWriter에서 안전하게 다룰 수 있는 범위 안에서만 허용한다.
    기본값은 ``IMAGE_COLUMN_WIDTH``/``IMAGE_ROW_HEIGHT``이지만 호출자가 직접
    값을 지정할 수 있다.
    """

    try:
        width = float(column_width)
        height = float(row_height)
    except (TypeError, ValueError) as exc:
        raise ValueError("열 너비와 행 높이는 숫자로 입력해 주세요.") from exc

    if not math.isfinite(width) or not math.isfinite(height):
        raise ValueError("열 너비와 행 높이는 유한한 숫자여야 합니다.")
    if not 0.1 <= width <= 255.0:
        raise ValueError("열 너비는 0.1 이상 255 이하로 입력해 주세요.")
    if not 1.0 <= height <= 409.0:
        raise ValueError("행 높이는 1 이상 409 이하로 입력해 주세요.")
    return width, height


def validate_output_filename(filename: str) -> str:
    """Windows에서도 안전하게 저장할 수 있는 Excel 파일명인지 확인한다."""

    name = filename.strip()
    if not name:
        raise ValueError("Excel 파일명을 입력해 주세요.")
    if name in {".", ".."} or INVALID_WINDOWS_FILENAME_RE.search(name):
        raise ValueError('파일명에는 \\ / : * ? " < > | 문자를 사용할 수 없습니다.')
    if name.endswith((" ", ".")):
        raise ValueError("파일명 끝에는 공백이나 마침표를 사용할 수 없습니다.")

    # Windows 예약 장치명은 확장자가 붙어 있어도 파일명으로 사용할 수 없다.
    first_part = name.split(".", 1)[0].rstrip(" .").upper()
    if first_part in WINDOWS_RESERVED_BASENAMES:
        raise ValueError(f"'{first_part}'은 Windows에서 사용할 수 없는 예약 파일명입니다.")
    return name


def _read_svg_source(path: Path) -> bytes | None:
    """확장자와 관계없이 실제 내용이 SVG/SVGZ이면 SVG 원문을 반환한다."""

    try:
        source = path.read_bytes()
    except OSError:
        return None

    if source.startswith(b"\x1f\x8b"):
        try:
            source = gzip.decompress(source)
        except (OSError, EOFError):
            return None

    try:
        root = ElementTree.fromstring(source)
    except (ElementTree.ParseError, ValueError):
        return None

    tag = root.tag if isinstance(root.tag, str) else ""
    local_name = tag.rsplit("}", 1)[-1].casefold()
    return source if local_name == "svg" else None


def _open_svg_source(source: bytes) -> Image.Image:
    try:
        import cairosvg
    except (ImportError, OSError) as exc:
        raise ImageConversionError(
            "SVG 변환기가 없습니다. 'python -m pip install CairoSVG'를 설치해 주세요."
        ) from exc

    try:
        png_bytes = cairosvg.svg2png(bytestring=source, unsafe=False)
        image = Image.open(BytesIO(png_bytes))
        image.load()
        return image
    except Exception as exc:
        raise ImageConversionError(f"SVG 변환 실패: {exc}") from exc


def _open_camera_raw(path: Path) -> Image.Image:
    try:
        import rawpy
    except (ImportError, OSError) as exc:
        raise ImageConversionError(
            "카메라 RAW 변환기가 없습니다. 'python -m pip install rawpy'를 설치해 주세요."
        ) from exc

    try:
        with rawpy.imread(str(path)) as raw:
            rgb = raw.postprocess(use_camera_wb=True, output_bps=8)
        return Image.fromarray(rgb)
    except Exception as exc:
        raise ImageConversionError(f"카메라 RAW 변환 실패: {exc}") from exc


def _open_image_any_format(path: Path) -> Image.Image:
    """확장자보다 파일 내용을 우선해 이미지를 연다."""

    suffix = path.suffix.casefold()
    try:
        image = Image.open(path)
        image.load()
        return image
    except (UnidentifiedImageError, OSError, ValueError) as pillow_error:
        svg_source = _read_svg_source(path)
        if svg_source is not None:
            return _open_svg_source(svg_source)
        try:
            return _open_camera_raw(path)
        except ImageConversionError as raw_error:
            if suffix in {".heic", ".heif", ".hif"} and not OPTIONAL_DECODERS["HEIC/HEIF"]:
                raise ImageConversionError(
                    "HEIC/HEIF 변환기가 없습니다. 'python -m pip install pillow-heif'를 설치해 주세요."
                ) from pillow_error
            raise ImageConversionError(
                f"지원되지 않거나 손상된 이미지입니다. Pillow 오류: {pillow_error}; "
                f"RAW 시도: {raw_error}"
            ) from pillow_error


def convert_to_lossless_png(source: Path, destination: Path) -> tuple[int, int, int, str]:
    """이미지를 리사이즈하지 않고 Excel 호환 PNG로 변환한다.

    반환값은 ``(가로 픽셀, 세로 픽셀, 프레임 수, 원본 형식)``이다.
    다중 페이지 TIFF/애니메이션 파일은 파일 1개당 셀 1개 원칙에 따라 첫 프레임을 쓴다.
    """

    if not source.is_file():
        raise ImageConversionError("파일이 없거나 일반 파일이 아닙니다.")

    image: Image.Image | None = None
    try:
        image = _open_image_any_format(source)
        source_format = image.format or source.suffix.lstrip(".").upper() or "UNKNOWN"
        frame_count = int(getattr(image, "n_frames", 1))
        if frame_count > 1:
            image.seek(0)

        oriented = ImageOps.exif_transpose(image)
        oriented.load()

        has_transparency = oriented.mode in {"RGBA", "LA"} or (
            oriented.mode == "P" and "transparency" in oriented.info
        )
        converted = oriented.convert("RGBA" if has_transparency else "RGB")

        save_options: dict[str, object] = {
            "format": "PNG",
            "compress_level": 6,
            "optimize": False,
        }
        icc_profile = oriented.info.get("icc_profile")
        if icc_profile:
            save_options["icc_profile"] = icc_profile

        converted.save(destination, **save_options)
        width, height = converted.size
        converted.close()
        return width, height, frame_count, source_format
    except ImageConversionError:
        raise
    except Exception as exc:
        raise ImageConversionError(f"이미지 변환 실패: {exc}") from exc
    finally:
        if image is not None:
            image.close()


def _merge_or_write_title(worksheet, first_col: int, last_col: int, text: str, cell_format) -> None:
    if first_col == last_col:
        worksheet.write_string(0, first_col, text, cell_format)
    else:
        worksheet.merge_range(0, first_col, 0, last_col, text, cell_format)


def _write_result_sheet(workbook, results: Sequence[ImageResult]) -> None:
    sheet = workbook.add_worksheet("처리 결과")
    header = workbook.add_format(
        {
            "font_name": EXCEL_FONT_NAME,
            "bold": True,
            "font_color": "#FFFFFF",
            "bg_color": "#4472C4",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
        }
    )
    normal = workbook.add_format(
        {"font_name": EXCEL_FONT_NAME, "border": 1, "valign": "top"}
    )
    ok_format = workbook.add_format(
        {
            "font_name": EXCEL_FONT_NAME,
            "border": 1,
            "font_color": "#008000",
            "align": "center",
        }
    )
    fail_format = workbook.add_format(
        {
            "font_name": EXCEL_FONT_NAME,
            "border": 1,
            "font_color": "#C00000",
            "bg_color": "#FCE4D6",
            "align": "center",
        }
    )

    headings = [
        "전체 순번",
        "행 그룹",
        "그룹 내 순번",
        "파일명",
        "확장자",
        "원본 크기(px)",
        "프레임 수",
        "원본 형식",
        "상태",
        "비고",
        "원본 경로",
    ]
    for col, value in enumerate(headings):
        sheet.write_string(0, col, value, header)

    for row, result in enumerate(results, start=1):
        values = [
            result.order,
            result.group,
            result.position_in_group,
            result.path.name,
            result.path.suffix,
            (
                f"{result.pixel_width} × {result.pixel_height}"
                if result.pixel_width is not None and result.pixel_height is not None
                else ""
            ),
            result.frame_count if result.frame_count is not None else "",
            result.source_format or "",
            "성공" if result.success else "실패",
            result.note,
            str(result.path),
        ]
        for col, value in enumerate(values):
            if col == 8:
                sheet.write(row, col, value, ok_format if result.success else fail_format)
            elif isinstance(value, int):
                sheet.write_number(row, col, value, normal)
            else:
                sheet.write_string(row, col, str(value), normal)

    sheet.set_row(0, 24)
    widths = [10, 12, 12, 28, 10, 18, 10, 12, 10, 45, 60]
    for col, width in enumerate(widths):
        sheet.set_column(col, col, width)
    sheet.freeze_panes(1, 0)
    sheet.autofilter(0, 0, max(len(results), 1), len(headings) - 1)
    sheet.hide_gridlines(2)


def verify_xlsx_package(path: Path, expected_image_anchors: int) -> None:
    """XLSX의 ZIP CRC, XML 문법, 필수 파일, 이미지 앵커 수를 검사한다."""

    required_parts = {
        "[Content_Types].xml",
        "_rels/.rels",
        "xl/workbook.xml",
        "xl/worksheets/sheet1.xml",
    }

    try:
        with zipfile.ZipFile(path, "r") as archive:
            names = set(archive.namelist())
            missing = required_parts - names
            if missing:
                raise RuntimeError(f"XLSX 필수 구성 파일 누락: {sorted(missing)}")

            bad_member = archive.testzip()
            if bad_member is not None:
                raise RuntimeError(f"XLSX ZIP CRC 오류: {bad_member}")

            for name in names:
                if name.endswith((".xml", ".rels")):
                    try:
                        ElementTree.fromstring(archive.read(name))
                    except ElementTree.ParseError as exc:
                        raise RuntimeError(f"XLSX XML 오류({name}): {exc}") from exc

            drawing_files = sorted(
                name
                for name in names
                if name.startswith("xl/drawings/drawing") and name.endswith(".xml")
            )
            blip_namespace = "{http://schemas.openxmlformats.org/drawingml/2006/main}blip"
            actual_image_anchors = 0
            for drawing_name in drawing_files:
                root = ElementTree.fromstring(archive.read(drawing_name))
                actual_image_anchors += sum(1 for _ in root.iter(blip_namespace))

            if actual_image_anchors != expected_image_anchors:
                raise RuntimeError(
                    "이미지 개체 수 검증 실패: "
                    f"예상 {expected_image_anchors}개, 실제 {actual_image_anchors}개"
                )
    except zipfile.BadZipFile as exc:
        raise RuntimeError("생성 파일이 올바른 XLSX ZIP 패키지가 아닙니다.") from exc


def create_excel(
    paths: Sequence[str | os.PathLike[str]],
    output_path: str | os.PathLike[str],
    *,
    overwrite: bool = False,
    progress: ProgressCallback | None = None,
    image_column_width: float = IMAGE_COLUMN_WIDTH,
    image_row_height: float = IMAGE_ROW_HEIGHT,
    manual_order: Sequence[str | os.PathLike[str]] | None = None,
    cancel_check: CancelCheck | None = None,
) -> GenerationSummary:
    """선택 이미지를 정렬해 XLSX를 만들고 검증한다.

    ``image_column_width``와 ``image_row_height``는 이미지 셀의 표시 크기다.
    이미지 원본 픽셀은 그대로 유지하고 XlsxWriter의 표시 배율만 바꾼다.
    """

    image_column_width, image_row_height = validate_image_cell_size(
        image_column_width, image_row_height
    )
    groups = organize_paths(paths, manual_order=manual_order)
    if not groups:
        raise ValueError("선택된 파일이 없습니다.")

    selected_count = sum(len(items) for _, items in groups)
    max_group_size = max(len(items) for _, items in groups)
    group_count = len(groups)

    if DATA_START_ROW + group_count > MAX_EXCEL_ROWS:
        raise ValueError("행 그룹 수가 Excel 최대 행 수를 초과합니다.")

    # 이미지 표의 첫 열(A열)은 파일명에서 추출한 행 그룹을 표시한다.
    # 이미지들은 B열부터 시작하고, 이미지 표와 파일명 검증표 사이에는 빈 열 1개를 둔다.
    validation_start_col = max_group_size + 2
    validation_last_col = validation_start_col + max_group_size
    if validation_last_col >= MAX_EXCEL_COLUMNS:
        raise ValueError(
            "한 행의 이미지 수가 너무 많아 이미지 표와 검증표를 한 시트에 함께 배치할 수 없습니다."
        )

    output = Path(output_path).expanduser().resolve()
    if output.suffix.casefold() != ".xlsx":
        output = output.with_suffix(".xlsx")
    validate_output_filename(output.name)
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists() and not overwrite:
        raise FileExistsError(f"이미 존재하는 파일입니다: {output}")

    temporary_output = output.with_name(f".{output.stem}.{uuid.uuid4().hex}.tmp.xlsx")
    results: list[ImageResult] = []
    inserted_count = 0

    try:
        with tempfile.TemporaryDirectory(prefix="excel_image_organizer_") as conversion_dir:
            conversion_root = Path(conversion_dir)
            with xlsxwriter.Workbook(str(temporary_output)) as workbook:
                sheet = workbook.add_worksheet("이미지 정리")
                sheet.set_tab_color("#4472C4")
                sheet.hide_gridlines(2)
                sheet.set_zoom(80)
                sheet.freeze_panes(DATA_START_ROW, 0)
                sheet.set_landscape()
                sheet.fit_to_pages(1, 0)

                title_format = workbook.add_format(
                    {
                        "font_name": EXCEL_FONT_NAME,
                        "bold": True,
                        "font_size": 12,
                        "font_color": "#FFFFFF",
                        "bg_color": "#1F4E78",
                        "border": 1,
                        "align": "center",
                        "valign": "vcenter",
                    }
                )
                header_format = workbook.add_format(
                    {
                        "font_name": EXCEL_FONT_NAME,
                        "bold": True,
                        "font_color": "#FFFFFF",
                        "bg_color": "#4472C4",
                        "border": 1,
                        "align": "center",
                        "valign": "vcenter",
                    }
                )
                image_cell_format = workbook.add_format(
                    {
                        "font_name": EXCEL_FONT_NAME,
                        "border": 1,
                        "align": "center",
                        "valign": "vcenter",
                        "font_size": 8,
                    }
                )
                group_format = workbook.add_format(
                    {
                        "font_name": EXCEL_FONT_NAME,
                        "bold": True,
                        "bg_color": "#D9EAF7",
                        "border": 1,
                        "align": "center",
                        "valign": "vcenter",
                    }
                )
                filename_format = workbook.add_format(
                    {
                        "font_name": EXCEL_FONT_NAME,
                        "border": 1,
                        "align": "center",
                        "valign": "vcenter",
                        "text_wrap": True,
                        "font_size": 8,
                    }
                )
                failed_filename_format = workbook.add_format(
                    {
                        "font_name": EXCEL_FONT_NAME,
                        "border": 1,
                        "align": "center",
                        "valign": "vcenter",
                        "text_wrap": True,
                        "font_size": 8,
                        "font_color": "#C00000",
                        "bg_color": "#FCE4D6",
                    }
                )
                failed_image_format = workbook.add_format(
                    {
                        "font_name": EXCEL_FONT_NAME,
                        "border": 1,
                        "align": "center",
                        "valign": "vcenter",
                        "text_wrap": True,
                        "font_size": 8,
                        "font_color": "#C00000",
                        "bg_color": "#FCE4D6",
                    }
                )

                sheet.set_row(0, TITLE_ROW_HEIGHT)
                sheet.set_row(1, HEADER_ROW_HEIGHT)
                sheet.set_column(0, 0, VALIDATION_GROUP_WIDTH)
                sheet.set_column(1, max_group_size, image_column_width)
                sheet.set_column(max_group_size + 1, max_group_size + 1, VALIDATION_GAP_WIDTH)
                sheet.set_column(
                    validation_start_col,
                    validation_start_col,
                    VALIDATION_GROUP_WIDTH,
                )
                sheet.set_column(
                    validation_start_col + 1,
                    validation_last_col,
                    VALIDATION_FILENAME_WIDTH,
                )

                _merge_or_write_title(
                    sheet,
                    0,
                    max_group_size,
                    "이미지 정렬표",
                    title_format,
                )
                _merge_or_write_title(
                    sheet,
                    validation_start_col,
                    validation_last_col,
                    "파일명 검증표 (실제 배치 순서)",
                    title_format,
                )

                sheet.write_string(1, 0, "행 그룹", header_format)
                for col in range(max_group_size):
                    sheet.write_string(1, col + 1, f"이미지 {col + 1}", header_format)
                sheet.write_string(1, validation_start_col, "행 그룹", header_format)
                for position in range(max_group_size):
                    sheet.write_string(
                        1,
                        validation_start_col + 1 + position,
                        f"파일명 {position + 1}",
                        header_format,
                    )

                target_width_px = float(excel_column_width_to_pixels(image_column_width))
                target_height_px = float(excel_row_height_to_pixels(image_row_height))
                completed = 0
                overall_order = 0

                for group_index, (group_name, items) in enumerate(groups):
                    if cancel_check is not None and cancel_check():
                        raise GenerationCancelled("사용자가 Excel 생성을 취소했습니다.")
                    excel_row = DATA_START_ROW + group_index
                    sheet.set_row(excel_row, image_row_height)
                    # 이미지 정렬표와 파일명 검증표 양쪽에 같은 행 그룹 값을 표시한다.
                    sheet.write_string(excel_row, 0, group_name, group_format)
                    sheet.write_string(excel_row, validation_start_col, group_name, group_format)

                    for empty_col in range(max_group_size):
                        sheet.write_blank(excel_row, empty_col + 1, None, image_cell_format)
                        sheet.write_blank(
                            excel_row,
                            validation_start_col + 1 + empty_col,
                            None,
                            filename_format,
                        )

                    for position, item in enumerate(items, start=1):
                        if cancel_check is not None and cancel_check():
                            raise GenerationCancelled("사용자가 Excel 생성을 취소했습니다.")
                        overall_order += 1
                        image_col = position
                        validation_col = validation_start_col + position
                        converted_path = conversion_root / f"{overall_order:08d}.png"
                        note = ""

                        try:
                            width, height, frame_count, source_format = convert_to_lossless_png(
                                item.path, converted_path
                            )
                            if width <= 0 or height <= 0:
                                raise ImageConversionError("이미지 픽셀 크기가 0입니다.")

                            options = {
                                "x_offset": 0,
                                "y_offset": 0,
                                "x_scale": target_width_px / width,
                                "y_scale": target_height_px / height,
                                "object_position": 1,
                                "description": item.path.name,
                            }
                            insert_result = sheet.insert_image(
                                excel_row,
                                image_col,
                                str(converted_path),
                                options,
                            )
                            if insert_result != 0:
                                raise RuntimeError(f"XlsxWriter 이미지 삽입 반환 코드: {insert_result}")

                            if frame_count > 1:
                                note = f"다중 프레임 {frame_count}개 중 첫 프레임 사용"
                            sheet.write_string(
                                excel_row,
                                validation_col,
                                item.path.name,
                                filename_format,
                            )
                            sheet.write_comment(
                                excel_row,
                                image_col,                                f"파일명: {item.path.name}\n그룹: {group_name}\n원본: {item.path}",
                                {"author": "이미지 정리 프로그램"},
                            )
                            inserted_count += 1
                            result = ImageResult(
                                order=overall_order,
                                group=group_name,
                                position_in_group=position,
                                path=item.path,
                                success=True,
                                pixel_width=width,
                                pixel_height=height,
                                frame_count=frame_count,
                                source_format=source_format,
                                note=note,
                            )
                        except Exception as exc:
                            error_text = str(exc).strip() or exc.__class__.__name__
                            sheet.write_string(
                                excel_row,
                                image_col,
                                f"삽입 실패\n{item.path.name}",
                                failed_image_format,
                            )
                            sheet.write_string(
                                excel_row,
                                validation_col,
                                f"[실패] {item.path.name}",
                                failed_filename_format,
                            )
                            result = ImageResult(
                                order=overall_order,
                                group=group_name,
                                position_in_group=position,
                                path=item.path,
                                success=False,
                                note=error_text,
                            )

                        results.append(result)
                        completed += 1
                        if progress is not None:
                            progress(completed, selected_count, item.path.name)

                _write_result_sheet(workbook, results)
                workbook.set_properties(
                    {
                        "title": "이미지 자동 정리 결과",
                        "subject": "파일명 자연 정렬에 따른 이미지 배치",
                        "author": "이미지 Excel 자동 정리 프로그램",
                        "comments": "원본 픽셀을 유지하고 표시 크기만 셀에 맞춤",
                    }
                )

        verify_xlsx_package(temporary_output, expected_image_anchors=inserted_count)
        os.replace(temporary_output, output)
    except Exception:
        try:
            temporary_output.unlink(missing_ok=True)
        except OSError:
            pass
        raise

    failed_count = selected_count - inserted_count
    return GenerationSummary(
        output_path=output,
        selected_count=selected_count,
        inserted_count=inserted_count,
        failed_count=failed_count,
        group_count=group_count,
        results=tuple(results),
    )


class PreviewDialog:
    """정렬 결과를 미리 보고 같은 행 그룹 안에서 순서를 미세 조정한다."""

    def __init__(self, app: "ImageOrganizerApp") -> None:
        self.app = app
        self.window = tk.Toplevel(app.root)
        self.window.title("정렬 결과 미리보기")
        self.window.geometry("820x560")
        self.window.minsize(650, 420)
        self.window.transient(app.root)

        groups = organize_paths(
            app.selected_paths,
            manual_order=app.manual_order if app.manual_order else None,
        )
        self.order: list[Path] = [item.path for _, items in groups for item in items]

        outer = ttk.Frame(self.window, padding=12)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=1)

        summary = ttk.Label(
            outer,
            text=(
                f"총 {len(self.order)}개 · 행 그룹 {len(groups)}개 · "
                f"셀 크기 {app.image_column_width.get()} × {app.image_row_height.get()}"
            ),
        )
        summary.grid(row=0, column=0, sticky="w", pady=(0, 8))

        tree_frame = ttk.Frame(outer)
        tree_frame.grid(row=1, column=0, sticky="nsew")
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(
            tree_frame,
            columns=("order", "group", "filename", "folder"),
            show="headings",
            selectmode="browse",
        )
        self.tree.heading("order", text="순서")
        self.tree.heading("group", text="행 그룹")
        self.tree.heading("filename", text="파일명")
        self.tree.heading("folder", text="폴더")
        self.tree.column("order", width=55, anchor="center", stretch=False)
        self.tree.column("group", width=90, anchor="center", stretch=False)
        self.tree.column("filename", width=300, anchor="w")
        self.tree.column("folder", width=260, anchor="w")
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        yscroll.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=yscroll.set)

        warnings = analyze_selection(
            app.selected_paths, check_missing=bool(app.check_missing.get())
        )
        warning_text = "검사 결과: 이상 없음" if not warnings else "검사 결과: " + " | ".join(warnings[:5])
        if len(warnings) > 5:
            warning_text += f" | 외 {len(warnings) - 5}건"
        ttk.Label(
            outer,
            text=warning_text,
            foreground="#B00020" if warnings else "#2E7D32",
            wraplength=780,
        ).grid(row=2, column=0, sticky="w", pady=(8, 0))

        buttons = ttk.Frame(outer)
        buttons.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(buttons, text="↑ 위로", command=lambda: self._move(-1)).pack(side="left")
        ttk.Button(buttons, text="↓ 아래로", command=lambda: self._move(1)).pack(
            side="left", padx=(6, 0)
        )
        ttk.Button(buttons, text="자연정렬 복원", command=self._reset_natural).pack(
            side="left", padx=(12, 0)
        )
        ttk.Label(
            buttons,
            text="※ 수동 이동은 같은 행 그룹 안에서만 가능합니다.",
            foreground="#666666",
        ).pack(side="left", padx=(12, 0))
        ttk.Button(buttons, text="취소", command=self.window.destroy).pack(side="right")
        ttk.Button(buttons, text="이 순서 적용", command=self._apply).pack(
            side="right", padx=(0, 6)
        )
        self._refresh_tree()

    def _refresh_tree(self, selected_path: Path | None = None) -> None:
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        for index, path in enumerate(self.order, start=1):
            group = extract_group(path.name)
            iid = str(index - 1)
            self.tree.insert(
                "",
                "end",
                iid=iid,
                values=(index, group, path.name, str(path.parent)),
            )
            if selected_path is not None and path == selected_path:
                self.tree.selection_set(iid)
                self.tree.focus(iid)
                self.tree.see(iid)

    def _move(self, delta: int) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        index = int(selection[0])
        target = index + delta
        if target < 0 or target >= len(self.order):
            return
        current_path = self.order[index]
        target_path = self.order[target]
        if normalize_filename_text(extract_group(current_path.name)).casefold() != normalize_filename_text(
            extract_group(target_path.name)
        ).casefold():
            messagebox.showinfo(
                APP_TITLE,
                "행 그룹이 다른 파일끼리는 순서를 바꿀 수 없습니다.",
                parent=self.window,
            )
            return
        self.order[index], self.order[target] = self.order[target], self.order[index]
        self._refresh_tree(selected_path=current_path)

    def _reset_natural(self) -> None:
        groups = organize_paths(self.app.selected_paths)
        self.order = [item.path for _, items in groups for item in items]
        self._refresh_tree()

    def _apply(self) -> None:
        self.app.manual_order = list(self.order)
        self.app._refresh_file_list(preserve_manual=True)
        self.app.status_text.set("미리보기에서 조정한 순서를 적용했습니다.")
        self.window.destroy()


class CompletionDialog:
    """생성 결과를 보여주고 Excel/저장 폴더를 바로 열 수 있게 한다."""

    def __init__(self, app: "ImageOrganizerApp", summary: GenerationSummary) -> None:
        self.app = app
        self.summary = summary
        window = tk.Toplevel(app.root)
        self.window = window
        window.title("Excel 생성 완료")
        window.geometry("560x310")
        window.resizable(False, False)
        window.transient(app.root)
        window.grab_set()

        outer = ttk.Frame(window, padding=18)
        outer.pack(fill="both", expand=True)
        title = "✓ Excel 생성 완료" if not summary.failed_count else "Excel 생성 완료 (일부 실패)"
        ttk.Label(outer, text=title, font=("맑은 고딕", 15, "bold")).pack(anchor="w")
        ttk.Label(
            outer,
            text=(
                f"총 선택: {summary.selected_count}개\n"
                f"정상 삽입: {summary.inserted_count}개\n"
                f"실패: {summary.failed_count}개\n"
                f"행 그룹: {summary.group_count}개"
            ),
            justify="left",
        ).pack(anchor="w", pady=(12, 8))
        ttk.Label(
            outer,
            text=f"저장 위치\n{summary.output_path}",
            wraplength=510,
            foreground="#444444",
        ).pack(anchor="w", pady=(0, 14))

        if summary.failed_count:
            ttk.Label(
                outer,
                text="실패 항목은 Excel의 '처리 결과' 시트에서 확인할 수 있습니다.",
                foreground="#B00020",
            ).pack(anchor="w", pady=(0, 10))

        buttons = ttk.Frame(outer)
        buttons.pack(fill="x", side="bottom")
        ttk.Button(buttons, text="Excel 열기", command=self._open_excel).pack(side="left")
        ttk.Button(buttons, text="폴더 열기", command=self._open_folder).pack(
            side="left", padx=(8, 0)
        )
        ttk.Button(buttons, text="닫기", command=window.destroy).pack(side="right")

    def _open_excel(self) -> None:
        self.app._open_in_system(self.summary.output_path)

    def _open_folder(self) -> None:
        self.app._open_in_system(self.summary.output_path.parent)


class ImageOrganizerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("900x760")
        self.root.minsize(760, 620)

        saved = load_settings()
        remember = bool(saved.get("remember_settings", True))
        width_value = IMAGE_COLUMN_WIDTH
        height_value = IMAGE_ROW_HEIGHT
        if remember:
            try:
                width_value, height_value = validate_image_cell_size(
                    saved.get("image_column_width", IMAGE_COLUMN_WIDTH),
                    saved.get("image_row_height", IMAGE_ROW_HEIGHT),
                )
            except ValueError:
                width_value, height_value = IMAGE_COLUMN_WIDTH, IMAGE_ROW_HEIGHT

        self.selected_paths: list[Path] = []
        self.manual_order: list[Path] = []
        self.last_source_dir = (
            str(saved.get("last_source_dir"))
            if remember and saved.get("last_source_dir")
            else str(Path.home())
        )
        saved_output_dir = str(saved.get("output_dir", "")) if remember else ""
        if saved_output_dir and not Path(saved_output_dir).is_dir():
            saved_output_dir = ""

        self.output_dir = tk.StringVar(value=saved_output_dir)
        self.output_name = tk.StringVar(
            value=f"이미지_정리_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        self.status_text = tk.StringVar(value="1단계: 이미지 파일을 선택해 주세요.")
        self.count_text = tk.StringVar(value="선택 파일: 0개")
        self.progress_value = tk.DoubleVar(value=0.0)
        self.image_column_width = tk.StringVar(value=f"{width_value:g}")
        self.image_row_height = tk.StringVar(value=f"{height_value:g}")
        self.remember_settings = tk.BooleanVar(value=remember)
        self.check_missing = tk.BooleanVar(
            value=bool(saved.get("check_missing", False)) if remember else False
        )
        self.size_preset = tk.StringVar(value="사용자 설정")
        for label, values in SIZE_PRESETS.items():
            if values and values == (width_value, height_value):
                self.size_preset.set(label)
                break

        self.events: "queue.Queue[tuple[str, object]]" = queue.Queue()
        self.cancel_event = threading.Event()
        self.is_generating = False

        self._build_ui()
        self._register_drag_and_drop()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(100, self._poll_worker_events)

    def _build_ui(self) -> None:
        outer = ttk.Frame(self.root, padding=16)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(2, weight=1)

        title = ttk.Label(outer, text=APP_TITLE, font=("맑은 고딕", 16, "bold"))
        title.grid(row=0, column=0, sticky="w", pady=(0, 8))

        instruction = ttk.Label(
            outer,
            text=(
                "이미지를 파일명 규칙으로 행 그룹/숫자 순 정렬합니다. "
                "미리보기에서 같은 행 그룹 안의 순서를 직접 조정할 수도 있습니다."
            ),
            wraplength=840,
        )
        instruction.grid(row=1, column=0, sticky="w", pady=(0, 8))

        file_frame = ttk.LabelFrame(outer, text="1. 이미지 파일 선택", padding=10)
        file_frame.grid(row=2, column=0, sticky="nsew")
        file_frame.columnconfigure(0, weight=1)
        file_frame.rowconfigure(1, weight=1)

        button_row = ttk.Frame(file_frame)
        button_row.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self.select_button = ttk.Button(
            button_row, text="파일 선택(여러 개 가능)", command=self._select_files
        )
        self.select_button.pack(side="left")
        self.remove_button = ttk.Button(
            button_row, text="선택 항목 제거", command=self._remove_selected
        )
        self.remove_button.pack(side="left", padx=(8, 0))
        self.clear_button = ttk.Button(button_row, text="전체 지우기", command=self._clear_files)
        self.clear_button.pack(side="left", padx=(8, 0))
        self.preview_button = ttk.Button(
            button_row, text="정렬 미리보기/순서 조정", command=self._show_preview
        )
        self.preview_button.pack(side="left", padx=(12, 0))
        ttk.Label(button_row, textvariable=self.count_text).pack(side="right")

        list_frame = ttk.Frame(file_frame)
        list_frame.grid(row=1, column=0, sticky="nsew")
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        self.file_list = tk.Listbox(list_frame, selectmode="extended", font=("맑은 고딕", 9))
        self.file_list.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.file_list.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.file_list.configure(yscrollcommand=scrollbar.set)
        self.drop_hint = ttk.Label(
            file_frame,
            text=(
                "탐색기 파일 드래그앤드롭 사용 가능"
                if DND_AVAILABLE
                else "드래그앤드롭: tkinterdnd2 설치 시 사용 가능 (기존 파일 선택은 정상 사용 가능)"
            ),
            foreground="#666666",
        )
        self.drop_hint.grid(row=2, column=0, sticky="w", pady=(6, 0))

        output_frame = ttk.LabelFrame(outer, text="2. 저장 및 옵션", padding=10)
        output_frame.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        output_frame.columnconfigure(1, weight=1)

        ttk.Label(output_frame, text="저장 폴더").grid(row=0, column=0, sticky="w")
        ttk.Entry(output_frame, textvariable=self.output_dir, state="readonly").grid(
            row=0, column=1, sticky="ew", padx=8
        )
        self.folder_button = ttk.Button(
            output_frame, text="폴더 지정", command=self._select_output_folder
        )
        self.folder_button.grid(row=0, column=2)

        ttk.Label(output_frame, text="엑셀 파일명").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.name_entry = ttk.Entry(output_frame, textvariable=self.output_name)
        self.name_entry.grid(row=1, column=1, columnspan=2, sticky="ew", padx=(8, 0), pady=(8, 0))

        ttk.Label(output_frame, text="크기 프리셋").grid(row=2, column=0, sticky="w", pady=(8, 0))
        self.preset_combo = ttk.Combobox(
            output_frame,
            textvariable=self.size_preset,
            values=list(SIZE_PRESETS),
            state="readonly",
            width=22,
        )
        self.preset_combo.grid(row=2, column=1, sticky="w", padx=8, pady=(8, 0))
        self.preset_combo.bind("<<ComboboxSelected>>", self._apply_size_preset)

        ttk.Label(output_frame, text="이미지 셀 크기").grid(
            row=3, column=0, sticky="w", pady=(8, 0)
        )
        size_frame = ttk.Frame(output_frame)
        size_frame.grid(row=3, column=1, columnspan=2, sticky="w", padx=(8, 0), pady=(8, 0))
        ttk.Label(size_frame, text="열 너비").pack(side="left")
        self.column_width_entry = ttk.Entry(
            size_frame, textvariable=self.image_column_width, width=8
        )
        self.column_width_entry.pack(side="left", padx=(6, 14))
        ttk.Label(size_frame, text="행 높이").pack(side="left")
        self.row_height_entry = ttk.Entry(
            size_frame, textvariable=self.image_row_height, width=8
        )
        self.row_height_entry.pack(side="left", padx=(6, 8))
        ttk.Label(
            size_frame,
            text=f"(기본값 {IMAGE_COLUMN_WIDTH:g} / {IMAGE_ROW_HEIGHT:g})",
            foreground="#666666",
        ).pack(side="left")

        option_frame = ttk.Frame(output_frame)
        option_frame.grid(row=4, column=1, columnspan=2, sticky="w", padx=8, pady=(8, 0))
        self.remember_check = ttk.Checkbutton(
            option_frame, text="마지막 설정 기억", variable=self.remember_settings
        )
        self.remember_check.pack(side="left")
        self.missing_check = ttk.Checkbutton(
            option_frame,
            text="파일 순번 누락 가능성 검사",
            variable=self.check_missing,
        )
        self.missing_check.pack(side="left", padx=(14, 0))

        action_frame = ttk.Frame(outer)
        action_frame.grid(row=4, column=0, sticky="ew", pady=(12, 0))
        action_frame.columnconfigure(0, weight=1)
        self.progress_bar = ttk.Progressbar(
            action_frame, variable=self.progress_value, maximum=100, mode="determinate"
        )
        self.progress_bar.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        self.cancel_button = ttk.Button(
            action_frame, text="취소", command=self._cancel_generation, state="disabled"
        )
        self.cancel_button.grid(row=0, column=1, padx=(0, 8))
        self.create_button = ttk.Button(
            action_frame, text="3. Excel로 내보내기", command=self._start_generation
        )
        self.create_button.grid(row=0, column=2)

        status = ttk.Label(outer, textvariable=self.status_text, wraplength=850)
        status.grid(row=5, column=0, sticky="w", pady=(8, 0))

        decoder_text = ", ".join(
            f"{name}: {'사용 가능' if enabled else '선택 설치 필요'}"
            for name, enabled in OPTIONAL_DECODERS.items()
        )
        ttk.Label(outer, text=f"추가 형식 — {decoder_text}", foreground="#666666").grid(
            row=6, column=0, sticky="w", pady=(5, 0)
        )

    def _register_drag_and_drop(self) -> None:
        if not DND_AVAILABLE or DND_FILES is None:
            return
        try:
            # tkinterdnd2가 제공하는 메서드는 런타임에 위젯에 주입된다.
            self.file_list.drop_target_register(DND_FILES)  # type: ignore[attr-defined]
            self.file_list.dnd_bind("<<Drop>>", self._on_drop_files)  # type: ignore[attr-defined]
        except (AttributeError, tk.TclError):
            self.drop_hint.configure(text="드래그앤드롭 초기화 실패 — 파일 선택 버튼을 사용해 주세요.")

    def _on_drop_files(self, event: object) -> None:
        raw_data = getattr(event, "data", "")
        if not raw_data:
            return
        try:
            values = self.root.tk.splitlist(raw_data)
        except tk.TclError:
            values = (str(raw_data),)
        self._add_paths(values)

    def _add_paths(self, values: Iterable[str | os.PathLike[str]]) -> None:
        existing = {os.path.normcase(str(path.resolve())) for path in self.selected_paths}
        added = 0
        for value in values:
            path = Path(value).expanduser().resolve()
            if not path.is_file():
                continue
            identity = os.path.normcase(str(path))
            if identity not in existing:
                self.selected_paths.append(path)
                existing.add(identity)
                added += 1
        if added:
            self.manual_order.clear()
            self._refresh_file_list()
            if not self.output_dir.get():
                self.status_text.set("2단계: 결과 Excel을 저장할 폴더를 지정해 주세요.")

    def _select_files(self) -> None:
        chosen = filedialog.askopenfilenames(
            parent=self.root,
            title="정리할 이미지 파일을 모두 선택하세요",
            initialdir=self.last_source_dir,
            filetypes=(("모든 파일(내용으로 이미지 판별)", "*"),),
        )
        if not chosen:
            return
        self.last_source_dir = str(Path(chosen[0]).parent)
        self._add_paths(chosen)

    def _remove_selected(self) -> None:
        indices = set(self.file_list.curselection())
        if not indices:
            return
        self.selected_paths = [
            path for index, path in enumerate(self.selected_paths) if index not in indices
        ]
        self.manual_order.clear()
        self._refresh_file_list()

    def _clear_files(self) -> None:
        self.selected_paths.clear()
        self.manual_order.clear()
        self._refresh_file_list()

    def _refresh_file_list(self, preserve_manual: bool = False) -> None:
        self.file_list.delete(0, tk.END)
        manual = self.manual_order if preserve_manual and self.manual_order else None
        preview_groups = organize_paths(self.selected_paths, manual_order=manual)
        sorted_items = [item for _, items in preview_groups for item in items]
        self.selected_paths = [item.path for item in sorted_items]
        if preserve_manual:
            self.manual_order = list(self.selected_paths)
        for item in sorted_items:
            self.file_list.insert(
                tk.END, f"[{item.group}]  {item.path.name}    —    {item.path.parent}"
            )
        self.count_text.set(f"선택 파일: {len(self.selected_paths)}개")

    def _show_preview(self) -> None:
        if not self.selected_paths:
            messagebox.showwarning(APP_TITLE, "먼저 이미지 파일을 선택해 주세요.", parent=self.root)
            return
        try:
            validate_image_cell_size(self.image_column_width.get(), self.image_row_height.get())
        except ValueError as exc:
            messagebox.showwarning(APP_TITLE, str(exc), parent=self.root)
            return
        PreviewDialog(self)

    def _select_output_folder(self) -> None:
        initial = self.output_dir.get() or self.last_source_dir
        chosen = filedialog.askdirectory(
            parent=self.root,
            title="결과 Excel을 저장할 폴더를 지정하세요",
            initialdir=initial,
            mustexist=True,
        )
        if chosen:
            self.output_dir.set(chosen)
            self.status_text.set("3단계: 미리보기로 정렬을 확인하거나 Excel로 내보내세요.")

    def _apply_size_preset(self, _event: object | None = None) -> None:
        values = SIZE_PRESETS.get(self.size_preset.get())
        if values is None:
            return
        width, height = values
        self.image_column_width.set(f"{width:g}")
        self.image_row_height.set(f"{height:g}")

    def _save_current_settings(self) -> None:
        if not self.remember_settings.get():
            try:
                save_settings({"remember_settings": False})
            except OSError:
                pass
            return
        try:
            width, height = validate_image_cell_size(
                self.image_column_width.get(), self.image_row_height.get()
            )
        except ValueError:
            width, height = IMAGE_COLUMN_WIDTH, IMAGE_ROW_HEIGHT
        settings = {
            "remember_settings": True,
            "image_column_width": width,
            "image_row_height": height,
            "last_source_dir": self.last_source_dir,
            "output_dir": self.output_dir.get(),
            "check_missing": bool(self.check_missing.get()),
        }
        try:
            save_settings(settings)
        except OSError:
            # 설정 저장 실패가 핵심 Excel 기능을 막아서는 안 된다.
            pass

    def _preflight_confirm(self) -> bool:
        warnings = analyze_selection(
            self.selected_paths, check_missing=bool(self.check_missing.get())
        )
        if not warnings:
            return True
        shown = "\n".join(f"• {warning}" for warning in warnings[:12])
        if len(warnings) > 12:
            shown += f"\n• 외 {len(warnings) - 12}건"
        return messagebox.askyesno(
            APP_TITLE,
            "사전 검사에서 확인할 항목이 있습니다.\n\n"
            f"{shown}\n\n그래도 Excel 생성을 계속할까요?",
            parent=self.root,
        )

    def _start_generation(self) -> None:
        if not self.selected_paths:
            messagebox.showwarning(APP_TITLE, "먼저 이미지 파일을 선택해 주세요.", parent=self.root)
            return
        if not self.output_dir.get():
            messagebox.showwarning(APP_TITLE, "저장 폴더를 지정해 주세요.", parent=self.root)
            return

        name = self.output_name.get().strip()
        try:
            name = validate_output_filename(name)
        except ValueError as exc:
            messagebox.showwarning(APP_TITLE, str(exc), parent=self.root)
            return
        if not name.casefold().endswith(".xlsx"):
            name += ".xlsx"
            self.output_name.set(name)

        try:
            image_column_width, image_row_height = validate_image_cell_size(
                self.image_column_width.get(), self.image_row_height.get()
            )
        except ValueError as exc:
            messagebox.showwarning(APP_TITLE, str(exc), parent=self.root)
            return
        self.image_column_width.set(f"{image_column_width:g}")
        self.image_row_height.set(f"{image_row_height:g}")

        if not self._preflight_confirm():
            return

        output = Path(self.output_dir.get()) / name
        overwrite = False
        if output.exists():
            overwrite = messagebox.askyesno(
                APP_TITLE,
                f"이미 같은 파일이 있습니다. 덮어쓸까요?\n\n{output}",
                parent=self.root,
            )
            if not overwrite:
                return

        self._save_current_settings()
        self._set_controls_enabled(False)
        self.cancel_event.clear()
        self.is_generating = True
        self.cancel_button.configure(state="normal")
        self.progress_value.set(0)
        self.status_text.set("이미지를 변환하고 Excel을 만드는 중입니다...")
        paths = tuple(self.selected_paths)
        manual_order = tuple(self.manual_order) if self.manual_order else None

        def report_progress(current: int, total: int, filename: str) -> None:
            self.events.put(("progress", (current, total, filename)))

        def worker() -> None:
            try:
                summary = create_excel(
                    paths,
                    output,
                    overwrite=overwrite,
                    progress=report_progress,
                    image_column_width=image_column_width,
                    image_row_height=image_row_height,
                    manual_order=manual_order,
                    cancel_check=self.cancel_event.is_set,
                )
                self.events.put(("done", summary))
            except GenerationCancelled as exc:
                self.events.put(("cancelled", exc))
            except Exception as exc:
                self.events.put(("error", exc))

        threading.Thread(target=worker, daemon=True).start()

    def _cancel_generation(self) -> None:
        if not self.is_generating:
            return
        self.cancel_event.set()
        self.cancel_button.configure(state="disabled")
        self.status_text.set("취소 요청됨 — 현재 파일 처리 후 안전하게 중단합니다...")

    def _finish_generation_state(self) -> None:
        self.is_generating = False
        self.cancel_button.configure(state="disabled")
        self._set_controls_enabled(True)

    def _poll_worker_events(self) -> None:
        try:
            while True:
                event, payload = self.events.get_nowait()
                if event == "progress":
                    current, total, filename = payload  # type: ignore[misc]
                    percent = (current / total * 100.0) if total else 0.0
                    self.progress_value.set(percent)
                    self.status_text.set(f"처리 중 {current}/{total}: {filename}")
                elif event == "done":
                    summary = payload
                    assert isinstance(summary, GenerationSummary)
                    self._finish_generation_state()
                    self.progress_value.set(100)
                    self.status_text.set(
                        f"완료: {summary.inserted_count}개 성공, {summary.failed_count}개 실패, "
                        f"{summary.group_count}개 행 그룹"
                    )
                    CompletionDialog(self, summary)
                elif event == "cancelled":
                    self._finish_generation_state()
                    self.progress_value.set(0)
                    self.status_text.set("사용자 요청으로 Excel 생성을 취소했습니다.")
                    messagebox.showinfo(
                        APP_TITLE,
                        "Excel 생성을 안전하게 취소했습니다. 임시 결과 파일은 정리됩니다.",
                        parent=self.root,
                    )
                elif event == "error":
                    self._finish_generation_state()
                    self.status_text.set("생성 실패 — 아래 오류를 확인해 주세요.")
                    messagebox.showerror(APP_TITLE, str(payload), parent=self.root)
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self._poll_worker_events)

    def _open_in_system(self, path: Path) -> None:
        try:
            if os.name != "nt" or not hasattr(os, "startfile"):
                raise RuntimeError("이 기능은 Windows 탐색기/Excel 환경에서 사용할 수 있습니다.")
            os.startfile(str(path))  # type: ignore[attr-defined]
        except (OSError, RuntimeError) as exc:
            messagebox.showerror(APP_TITLE, f"열 수 없습니다.\n{exc}", parent=self.root)

    def _set_controls_enabled(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        readonly_state = "readonly" if enabled else "disabled"
        for widget in (
            self.select_button,
            self.remove_button,
            self.clear_button,
            self.preview_button,
            self.folder_button,
            self.name_entry,
            self.column_width_entry,
            self.row_height_entry,
            self.remember_check,
            self.missing_check,
            self.create_button,
        ):
            widget.configure(state=state)
        self.preset_combo.configure(state=readonly_state)

    def _on_close(self) -> None:
        if self.is_generating:
            if not messagebox.askyesno(
                APP_TITLE,
                "Excel 생성 중입니다. 작업을 취소하고 프로그램을 종료할까요?",
                parent=self.root,
            ):
                return
            self.cancel_event.set()
        self._save_current_settings()
        self.root.destroy()

def main() -> None:
    root = TkinterDnD.Tk() if DND_AVAILABLE and TkinterDnD is not None else tk.Tk()
    try:
        ttk.Style(root).theme_use("vista")
    except tk.TclError:
        pass
    ImageOrganizerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()