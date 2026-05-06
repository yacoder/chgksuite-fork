"""Split a multi-handout .hndt file and maximize rows per one-page handout."""

from __future__ import annotations

import math
import os
import subprocess
import sys
import tempfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from threading import Lock
from types import SimpleNamespace
from typing import Callable

from pypdf import PdfReader
from pypdf.generic import ContentStream, IndirectObject

from chgksuite.handouter.installer import get_tectonic_path, install_tectonic
from chgksuite.handouter.runner import HandoutGenerator, get_num_teams
from chgksuite.handouter.utils import compress_pdf

RESERVED_WORDS = {
    "image",
    "for_question",
    "columns",
    "rows",
    "resize_image",
    "font_size",
    "font_family",
    "no_center",
    "raw_tex",
    "color",
    "handouts_per_team",
    "grouping",
    "rotate",
    "tikz_mm",
    "hspace",
    "vspace",
    "max_width",
}

TECTONIC_LOCK = Lock()
TECTONIC_PATH: str | None = None
RENDERER_DEFAULTS = {
    "debug": False,
    "compress_pdf": "off",
    "optimize_images": "on",
    "font": None,
    "font_size": 14,
    "paperwidth": 210,
    "paperheight": 297,
    "margin_top": 5,
    "margin_bottom": 5,
    "margin_left": 5,
    "margin_right": 5,
    "boxwidth": None,
    "boxwidthinner": None,
    "tikz_mm": None,
    "add_n_teams": "off",
    "tectonic_package_regex": None,
}


@dataclass(frozen=True)
class HandoutBlock:
    ordinal: int
    text: str
    meta: dict[str, str]


@dataclass
class ProbeResult:
    rows: int
    ok: bool
    pages: int | None
    pdf_path: Path | None
    bottom_space_mm: float | None
    stdout: str
    stderr: str

    @property
    def fits_one_page(self) -> bool:
        return self.ok and self.pages is not None and self.pages <= 1


@dataclass
class BlockFitResult:
    ordinal: int
    output_path: Path
    question: str
    rows: int | None
    columns: int | None
    n_per_team: int | None
    resize_image: float | None
    logs: list[str]
    error: str | None = None


@dataclass(frozen=True)
class ImageResizeConfig:
    enabled: bool
    bottom_space_row_ratio: float
    shrink_percent: float
    min_resize_image: float
    refine_iterations: int = 8


def split_blocks(contents: str) -> list[str]:
    blocks: list[list[str]] = []
    current: list[str] = []
    for line in contents.splitlines():
        if line.strip() == "---":
            blocks.append(current)
            current = []
        else:
            current.append(line)
    blocks.append(current)
    return ["\n".join(block).strip("\n") for block in blocks if "\n".join(block).strip()]


def metadata_line(line: str) -> tuple[str, str] | None:
    key, sep, value = line.partition(":")
    key = key.strip()
    if sep and key in RESERVED_WORDS:
        return key, value.strip()
    return None


def parse_blocks(contents: str) -> list[HandoutBlock]:
    result = []
    for ordinal, block in enumerate(split_blocks(contents), start=1):
        meta: dict[str, str] = {}
        for line in block.splitlines():
            parsed = metadata_line(line)
            if parsed:
                key, value = parsed
                meta[key] = value
        result.append(HandoutBlock(ordinal=ordinal, text=block, meta=meta))
    return result


def parse_positive_int(value: str | None, key: str, default: int | None = None) -> int:
    if value is None:
        if default is None:
            raise ValueError(f"missing required `{key}`")
        value_int = default
    else:
        value_int = int(value)
    if value_int <= 0:
        raise ValueError(f"`{key}` must be positive, got {value_int}")
    return value_int


def parse_positive_float(
    value: str | None, key: str, default: float | None = None
) -> float:
    if value is None:
        if default is None:
            raise ValueError(f"missing required `{key}`")
        value_float = default
    else:
        value_float = float(value)
    if value_float <= 0:
        raise ValueError(f"`{key}` must be positive, got {value_float}")
    return value_float


def format_float(value: float) -> str:
    floored = math.floor(value * 100 + 1e-9) / 100
    return f"{floored:.2f}".rstrip("0").rstrip(".")


def valid_row_step(columns: int, n_per_team: int) -> int:
    return n_per_team // math.gcd(columns, n_per_team)


def block_max_width(block: HandoutBlock) -> float:
    max_width = parse_positive_float(block.meta.get("max_width"), "max_width", 1.0)
    if max_width > 1:
        raise ValueError(f"`max_width` must be <= 1, got {max_width}")
    return max_width


def max_width_column_multiplier(block: HandoutBlock) -> int:
    return max(1, math.floor(1.0 / block_max_width(block) + 1e-9))


def split_fit_columns(block: HandoutBlock) -> int:
    columns = parse_positive_int(block.meta.get("columns"), "columns")
    return columns * max_width_column_multiplier(block)


def upsert_metadata(block: HandoutBlock, updates: dict[str, str | None]) -> str:
    lines = block.text.splitlines()
    output: list[str] = []
    updated = set()

    for line in lines:
        parsed = metadata_line(line)
        if parsed and parsed[0] in updates:
            key = parsed[0]
            if key not in updated:
                if updates[key] is not None:
                    output.append(f"{key}: {updates[key]}")
                updated.add(key)
            continue
        output.append(line)

    missing = [
        key for key, value in updates.items() if key not in updated and value is not None
    ]
    if missing:
        insert_at = 0
        for idx, line in enumerate(output):
            if metadata_line(line):
                insert_at = idx + 1
        for key in missing:
            output.insert(insert_at, f"{key}: {updates[key]}")
            insert_at += 1

    return "\n".join(output).rstrip() + "\n"


def source_relative_image_update(
    block: HandoutBlock, source_dir: Path, output_dir: Path
) -> str | None:
    image = block.meta.get("image")
    if not image:
        return None
    image_path = Path(image)
    if image_path.is_absolute() or source_dir.resolve() == output_dir.resolve():
        return None
    return os.path.relpath(source_dir / image_path, output_dir)


def write_handout(
    block: HandoutBlock,
    output_path: Path,
    rows: int,
    source_dir: Path,
    resize_image: float | None = None,
) -> None:
    updates: dict[str, str | None] = {"rows": str(rows)}
    multiplier = max_width_column_multiplier(block)
    if multiplier > 1:
        columns = parse_positive_int(block.meta.get("columns"), "columns")
        updates["columns"] = str(columns * multiplier)
        updates["max_width"] = None
    if resize_image is not None:
        updates["resize_image"] = format_float(resize_image)
    image_update = source_relative_image_update(block, source_dir, output_path.parent)
    if image_update is not None:
        updates["image"] = image_update
    output_path.write_text(upsert_metadata(block, updates), encoding="utf8")


def build_renderer_args(hndt_path: Path, args) -> SimpleNamespace:
    values = {
        key: getattr(args, key, default) for key, default in RENDERER_DEFAULTS.items()
    }
    values["filename"] = str(hndt_path)
    values["language"] = args.language
    return SimpleNamespace(**values)


def cached_tectonic_path(renderer_args: SimpleNamespace) -> str | None:
    global TECTONIC_PATH
    if TECTONIC_PATH:
        return TECTONIC_PATH
    with TECTONIC_LOCK:
        if TECTONIC_PATH:
            return TECTONIC_PATH
        tectonic_path = get_tectonic_path()
        if not tectonic_path:
            install_tectonic(renderer_args)
            tectonic_path = get_tectonic_path()
        TECTONIC_PATH = tectonic_path
        return TECTONIC_PATH


IDENTITY_MATRIX = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
PATH_PAINT_OPERATORS = {b"S", b"s", b"f", b"F", b"f*", b"B", b"B*", b"b", b"b*"}
TEXT_SHOW_OPERATORS = {b"Tj", b"TJ", b"'", b'"'}


def pdf_deref(obj):
    return obj.get_object() if isinstance(obj, IndirectObject) else obj


def pdf_matrix(values) -> tuple[float, float, float, float, float, float]:
    return tuple(float(value) for value in values)  # type: ignore[return-value]


def pdf_matrix_multiply(left, right):
    return (
        left[0] * right[0] + left[2] * right[1],
        left[1] * right[0] + left[3] * right[1],
        left[0] * right[2] + left[2] * right[3],
        left[1] * right[2] + left[3] * right[3],
        left[0] * right[4] + left[2] * right[5] + left[4],
        left[1] * right[4] + left[3] * right[5] + left[5],
    )


def pdf_transform_point(matrix, x, y):
    return (
        matrix[0] * x + matrix[2] * y + matrix[4],
        matrix[1] * x + matrix[3] * y + matrix[5],
    )


def pdf_expand_bbox(bbox, points):
    if not points:
        return bbox
    xs = [point[0] for point in points]
    ys = [point[1] for point in points]
    points_bbox = (min(xs), min(ys), max(xs), max(ys))
    if bbox is None:
        return points_bbox
    return (
        min(bbox[0], points_bbox[0]),
        min(bbox[1], points_bbox[1]),
        max(bbox[2], points_bbox[2]),
        max(bbox[3], points_bbox[3]),
    )


def pdf_text_length(text_obj) -> int:
    if isinstance(text_obj, list):
        return sum(len(item) for item in text_obj if isinstance(item, str))
    if isinstance(text_obj, str):
        return len(text_obj)
    return 1


def pdf_xobject_bbox(name, resources, pdf, ctm, seen_forms):
    resources = pdf_deref(resources)
    xobjects = pdf_deref(resources.get("/XObject")) if resources else None
    if not xobjects or name not in xobjects:
        return None

    ref = xobjects[name]
    obj = pdf_deref(ref)
    subtype = obj.get("/Subtype")
    if subtype == "/Image":
        return pdf_expand_bbox(
            None,
            [
                pdf_transform_point(ctm, 0, 0),
                pdf_transform_point(ctm, 1, 0),
                pdf_transform_point(ctm, 0, 1),
                pdf_transform_point(ctm, 1, 1),
            ],
        )

    if subtype != "/Form":
        return None

    key = (ref.idnum, ref.generation) if isinstance(ref, IndirectObject) else id(obj)
    if key in seen_forms:
        return None
    seen_forms.add(key)

    form_matrix = pdf_deref(obj.get("/Matrix")) or IDENTITY_MATRIX
    form_ctm = pdf_matrix_multiply(ctm, pdf_matrix(form_matrix))
    return pdf_content_bbox(
        obj,
        pdf,
        obj.get("/Resources") or resources,
        ctm=form_ctm,
        seen_forms=seen_forms,
    )


def pdf_content_bbox(
    content,
    pdf,
    resources,
    ctm=IDENTITY_MATRIX,
    seen_forms=None,
):
    if seen_forms is None:
        seen_forms = set()

    try:
        stream = ContentStream(content, pdf)
    except Exception:
        return None

    graphics_stack = []
    current_path = []
    bbox = None
    text_matrix = IDENTITY_MATRIX
    text_line_matrix = IDENTITY_MATRIX
    font_size = 0.0
    leading = 0.0

    for operands, operator in stream.operations:
        if operator == b"q":
            graphics_stack.append(ctm)
        elif operator == b"Q":
            ctm = graphics_stack.pop() if graphics_stack else IDENTITY_MATRIX
        elif operator == b"cm" and len(operands) >= 6:
            ctm = pdf_matrix_multiply(ctm, pdf_matrix(operands[:6]))
        elif operator in (b"m", b"l") and len(operands) >= 2:
            current_path.append(
                pdf_transform_point(ctm, float(operands[0]), float(operands[1]))
            )
        elif operator == b"c" and len(operands) >= 6:
            for index in (0, 2, 4):
                current_path.append(
                    pdf_transform_point(
                        ctm, float(operands[index]), float(operands[index + 1])
                    )
                )
        elif operator in (b"v", b"y") and len(operands) >= 4:
            for index in (0, 2):
                current_path.append(
                    pdf_transform_point(
                        ctm, float(operands[index]), float(operands[index + 1])
                    )
                )
        elif operator == b"re" and len(operands) >= 4:
            x, y, width, height = (float(value) for value in operands[:4])
            current_path.extend(
                [
                    pdf_transform_point(ctm, x, y),
                    pdf_transform_point(ctm, x + width, y),
                    pdf_transform_point(ctm, x, y + height),
                    pdf_transform_point(ctm, x + width, y + height),
                ]
            )
        elif operator in PATH_PAINT_OPERATORS:
            bbox = pdf_expand_bbox(bbox, current_path)
            current_path = []
        elif operator == b"n":
            current_path = []
        elif operator == b"BT":
            text_matrix = IDENTITY_MATRIX
            text_line_matrix = IDENTITY_MATRIX
        elif operator == b"Tf" and len(operands) >= 2:
            font_size = float(operands[1])
        elif operator == b"TL" and operands:
            leading = float(operands[0])
        elif operator in (b"Td", b"TD") and len(operands) >= 2:
            if operator == b"TD":
                leading = -float(operands[1])
            translation = (1.0, 0.0, 0.0, 1.0, float(operands[0]), float(operands[1]))
            text_line_matrix = pdf_matrix_multiply(text_line_matrix, translation)
            text_matrix = text_line_matrix
        elif operator == b"Tm" and len(operands) >= 6:
            text_matrix = pdf_matrix(operands[:6])
            text_line_matrix = text_matrix
        elif operator == b"T*":
            translation = (1.0, 0.0, 0.0, 1.0, 0.0, -leading)
            text_line_matrix = pdf_matrix_multiply(text_line_matrix, translation)
            text_matrix = text_line_matrix
        elif operator in TEXT_SHOW_OPERATORS:
            text_obj = operands[-1] if operands else ""
            width = max(font_size * 0.5 * pdf_text_length(text_obj), font_size * 0.5)
            text_ctm = pdf_matrix_multiply(ctm, text_matrix)
            bbox = pdf_expand_bbox(
                bbox,
                [
                    pdf_transform_point(text_ctm, 0.0, -0.3 * font_size),
                    pdf_transform_point(text_ctm, width, font_size),
                ],
            )
        elif operator == b"Do" and operands:
            xobject_bbox = pdf_xobject_bbox(operands[0], resources, pdf, ctm, seen_forms)
            if xobject_bbox is not None:
                bbox = pdf_expand_bbox(
                    bbox,
                    [
                        (xobject_bbox[0], xobject_bbox[1]),
                        (xobject_bbox[2], xobject_bbox[3]),
                    ],
                )

    return bbox


def pdf_bottom_space_mm(pdf_path: Path) -> float | None:
    reader = PdfReader(str(pdf_path))
    if not reader.pages:
        return None

    page = reader.pages[0]
    bbox = pdf_content_bbox(page.get_contents(), reader, page.get("/Resources"))
    if bbox is None:
        return None
    return max(0.0, bbox[1]) * 25.4 / 72.0


def run_hndt2pdf(
    hndt_path: Path,
    args,
    output_pdf_path: Path | None = None,
    compress_output: bool = False,
) -> ProbeResult:
    renderer_args = build_renderer_args(hndt_path, args)
    generator = HandoutGenerator(renderer_args)
    tex_contents = generator.generate()
    file_dir = hndt_path.parent
    base_name = hndt_path.stem
    add_n_teams = getattr(renderer_args, "add_n_teams", "off") == "on"
    num_teams = get_num_teams(str(hndt_path)) if add_n_teams else None
    if output_pdf_path is None:
        if num_teams is not None:
            pdf_base_name = f"{base_name}_{num_teams}teams_{args.language}"
        else:
            pdf_base_name = f"{base_name}_{args.language}"
        tex_path = file_dir / f"{pdf_base_name}.tex"
        pdf_path = tex_path.with_suffix(".pdf")
    else:
        pdf_path = output_pdf_path
        tex_path = output_pdf_path.with_suffix(".tex")
    tex_path.write_text(tex_contents, encoding="utf8")

    tectonic_path = cached_tectonic_path(renderer_args)
    if not tectonic_path:
        return ProbeResult(
            rows=-1,
            ok=False,
            pages=None,
            pdf_path=None,
            bottom_space_mm=None,
            stdout="",
            stderr="tectonic could not be found or installed",
        )

    proc = subprocess.run(
        [tectonic_path, tex_path.name],
        cwd=file_dir,
        text=True,
        capture_output=True,
        check=False,
    )
    stdout = proc.stdout or ""
    stderr = proc.stderr or ""
    for tmp in generator._temp_files:
        try:
            os.remove(tmp)
        except OSError:
            pass
    if not getattr(renderer_args, "debug", False):
        try:
            tex_path.unlink()
        except OSError:
            pass
    if proc.returncode != 0:
        return ProbeResult(
            rows=-1,
            ok=False,
            pages=None,
            pdf_path=pdf_path,
            bottom_space_mm=None,
            stdout=stdout,
            stderr=stderr,
        )

    try:
        pages = len(PdfReader(str(pdf_path)).pages)
        bottom_space = pdf_bottom_space_mm(pdf_path) if pages == 1 else None
        if compress_output:
            compress_pdf(str(pdf_path))
    except Exception as exc:
        return ProbeResult(
            rows=-1,
            ok=False,
            pages=None,
            pdf_path=pdf_path,
            bottom_space_mm=None,
            stdout=stdout,
            stderr=f"{stderr}\nCould not read {pdf_path}: {exc}".strip(),
        )

    return ProbeResult(
        rows=-1,
        ok=True,
        pages=pages,
        pdf_path=pdf_path,
        bottom_space_mm=bottom_space,
        stdout=stdout,
        stderr=stderr,
    )


def cleanup_pdf(result: ProbeResult) -> None:
    if result.pdf_path and result.pdf_path.exists():
        result.pdf_path.unlink()


def probe_rows(
    block: HandoutBlock,
    output_path: Path,
    rows: int,
    source_dir: Path,
    args,
    verbose: bool,
    resize_image: float | None = None,
    log: Callable[[str], None] | None = None,
) -> ProbeResult:
    write_handout(block, output_path, rows, source_dir, resize_image)
    result = run_hndt2pdf(output_path, args)
    result.rows = rows
    pages = "compile failed" if result.pages is None else f"{result.pages} page(s)"
    if verbose:
        suffix = ""
        if resize_image is not None:
            suffix += f", resize_image={format_float(resize_image)}"
        if result.bottom_space_mm is not None:
            suffix += f", bottom_space={result.bottom_space_mm:.1f}mm"
        message = f"  rows={rows}: {pages}{suffix}"
        if log:
            log(message)
        else:
            print(message)
    cleanup_pdf(result)
    return result


def best_rows_for_block(
    block: HandoutBlock,
    output_path: Path,
    source_dir: Path,
    args,
    max_rows: int,
    verbose: bool,
    resize_image: float | None = None,
    log: Callable[[str], None] | None = None,
) -> int:
    columns = split_fit_columns(block)
    n_per_team = parse_positive_int(
        block.meta.get("handouts_per_team"), "handouts_per_team", default=3
    )
    step = valid_row_step(columns, n_per_team)
    max_k = max_rows // step
    if max_k < 1:
        raise ValueError(
            f"no valid row count <= {max_rows}: columns={columns}, "
            f"handouts_per_team={n_per_team}, row step={step}"
        )

    cache: dict[int, ProbeResult] = {}

    def probe(k: int) -> ProbeResult:
        rows = k * step
        if rows not in cache:
            cache[rows] = probe_rows(
                block,
                output_path,
                rows,
                source_dir,
                args,
                verbose,
                resize_image,
                log,
            )
        return cache[rows]

    low = 0
    high = 1
    first_failure: int | None = None
    while high <= max_k:
        result = probe(high)
        if result.fits_one_page:
            low = high
            high *= 2
        else:
            first_failure = high
            break

    if low == 0:
        result = cache[step]
        details = result.stderr or result.stdout or "no compiler output"
        raise RuntimeError(
            f"minimum valid rows={step} does not fit or does not compile:\n{details}"
        )

    upper = first_failure or (max_k + 1)
    while low + 1 < upper:
        mid = (low + upper) // 2
        result = probe(mid)
        if result.fits_one_page:
            low = mid
        else:
            upper = mid

    best = low * step
    write_handout(block, output_path, best, source_dir, resize_image)
    return best


def block_resize_image(block: HandoutBlock) -> float:
    return parse_positive_float(block.meta.get("resize_image"), "resize_image", 1.0)


def resize_update_for_block(
    block: HandoutBlock, resize_image: float | None
) -> float | None:
    if resize_image is None or "image" not in block.meta:
        return None
    if "resize_image" in block.meta or not math.isclose(
        resize_image, 1.0, rel_tol=0, abs_tol=0.0001
    ):
        return resize_image
    return None


def probe_bottom_space(
    block: HandoutBlock,
    output_path: Path,
    rows: int,
    resize_image: float,
    source_dir: Path,
    args,
    verbose: bool,
    log: Callable[[str], None] | None,
) -> float | None:
    result = probe_rows(
        block=block,
        output_path=output_path,
        rows=rows,
        source_dir=source_dir,
        args=args,
        verbose=verbose,
        resize_image=resize_update_for_block(block, resize_image),
        log=log,
    )
    if not result.fits_one_page:
        return None
    return result.bottom_space_mm


def bottom_space_threshold_mm(
    bottom_space_mm: float,
    rows: int,
    args,
    ratio: float,
) -> float:
    available_height = (
        getattr(args, "paperheight", RENDERER_DEFAULTS["paperheight"])
        - getattr(args, "margin_top", RENDERER_DEFAULTS["margin_top"])
        - getattr(args, "margin_bottom", RENDERER_DEFAULTS["margin_bottom"])
    )
    row_height = max(0.0, available_height - bottom_space_mm) / rows
    return ratio * row_height


def max_resize_for_rows(
    block: HandoutBlock,
    output_path: Path,
    rows: int,
    low_resize: float,
    high_resize: float,
    source_dir: Path,
    args,
    config: ImageResizeConfig,
    verbose: bool,
    log: Callable[[str], None] | None,
) -> float:
    best = low_resize
    low = low_resize
    high = high_resize
    for _ in range(config.refine_iterations):
        mid = (low + high) / 2
        result = probe_rows(
            block=block,
            output_path=output_path,
            rows=rows,
            source_dir=source_dir,
            args=args,
            verbose=verbose,
            resize_image=resize_update_for_block(block, mid),
            log=log,
        )
        if result.fits_one_page:
            best = mid
            low = mid
        else:
            high = mid
    return best


def best_rows_at_resize(
    block: HandoutBlock,
    output_path: Path,
    resize_image: float,
    source_dir: Path,
    args,
    max_rows: int,
    verbose: bool,
    log: Callable[[str], None] | None,
) -> int:
    return best_rows_for_block(
        block=block,
        output_path=output_path,
        source_dir=source_dir,
        args=args,
        max_rows=max_rows,
        verbose=verbose,
        resize_image=resize_update_for_block(block, resize_image),
        log=log,
    )


def fit_rows_and_resize(
    block: HandoutBlock,
    output_path: Path,
    source_dir: Path,
    args,
    max_rows: int,
    resize_config: ImageResizeConfig,
    verbose: bool,
    log: Callable[[str], None] | None,
) -> tuple[int, float | None]:
    current_resize = block_resize_image(block)
    current_rows = best_rows_at_resize(
        block,
        output_path,
        current_resize,
        source_dir,
        args,
        max_rows,
        verbose,
        log,
    )
    if not resize_config.enabled or "image" not in block.meta:
        return current_rows, resize_update_for_block(block, current_resize)

    shrink_factor = 1 - resize_config.shrink_percent / 100
    while True:
        bottom_space = probe_bottom_space(
            block,
            output_path,
            current_rows,
            current_resize,
            source_dir,
            args,
            verbose,
            log,
        )
        if (
            bottom_space is None
            or current_resize <= resize_config.min_resize_image
        ):
            break
        threshold = bottom_space_threshold_mm(
            bottom_space,
            current_rows,
            args,
            resize_config.bottom_space_row_ratio,
        )
        if bottom_space <= threshold:
            break

        trial_resize = current_resize
        improved: tuple[float, int] | None = None
        while trial_resize > resize_config.min_resize_image:
            trial_resize = max(
                resize_config.min_resize_image, trial_resize * shrink_factor
            )
            trial_rows = best_rows_at_resize(
                block,
                output_path,
                trial_resize,
                source_dir,
                args,
                max_rows,
                verbose,
                log,
            )
            if trial_rows > current_rows:
                improved = (trial_resize, trial_rows)
                break
            if math.isclose(
                trial_resize,
                resize_config.min_resize_image,
                rel_tol=0,
                abs_tol=0.0001,
            ):
                break

        if improved is None:
            break

        low_resize, improved_rows = improved
        expanded_resize = max_resize_for_rows(
            block=block,
            output_path=output_path,
            rows=improved_rows,
            low_resize=low_resize,
            high_resize=current_resize,
            source_dir=source_dir,
            args=args,
            config=resize_config,
            verbose=verbose,
            log=log,
        )
        if log:
            log(
                "  image resize: "
                f"{format_float(current_resize)} -> {format_float(expanded_resize)}, "
                f"rows {current_rows} -> {improved_rows}"
            )
        current_resize = expanded_resize
        current_rows = improved_rows

    final_resize = resize_update_for_block(block, current_resize)
    write_handout(block, output_path, current_rows, source_dir, final_resize)
    return current_rows, final_resize


def output_path_for_block(
    block: HandoutBlock, source_stem: str, output_dir: Path, used: set[str]
) -> Path:
    question = block.meta.get("for_question")
    if question is None:
        suffix = f"{block.ordinal:02d}"
    else:
        suffix = f"q{int(question.split()[0]):02d}"
    candidate = f"{source_stem}_{suffix}.hndt"
    if candidate not in used:
        used.add(candidate)
        return output_dir / candidate

    counter = 2
    while True:
        candidate = f"{source_stem}_{suffix}_{counter}.hndt"
        if candidate not in used:
            used.add(candidate)
            return output_dir / candidate
        counter += 1


def compile_final_pdf(
    output_path: Path,
    args,
    expected_rows: int,
) -> None:
    result = run_hndt2pdf(
        output_path,
        args,
        compress_output=getattr(args, "compress_pdf", "off") == "on",
    )
    result.rows = expected_rows
    if not result.fits_one_page:
        details = result.stderr or result.stdout or "no compiler output"
        raise RuntimeError(f"final compile failed for {output_path}:\n{details}")


def fit_block(
    block: HandoutBlock,
    output_path: Path,
    source_dir: Path,
    args,
    max_rows: int,
    resize_config: ImageResizeConfig,
    keep_pdfs: bool,
    verbose: bool,
) -> BlockFitResult:
    logs: list[str] = []
    question = block.meta.get("for_question", str(block.ordinal))
    try:
        rows, resize_image = fit_rows_and_resize(
            block=block,
            output_path=output_path,
            source_dir=source_dir,
            args=args,
            max_rows=max_rows,
            resize_config=resize_config,
            verbose=verbose,
            log=logs.append,
        )
        columns = split_fit_columns(block)
        n_per_team = parse_positive_int(
            block.meta.get("handouts_per_team"),
            "handouts_per_team",
            default=3,
        )
        if keep_pdfs:
            compile_final_pdf(output_path, args, rows)
        return BlockFitResult(
            ordinal=block.ordinal,
            output_path=output_path,
            question=question,
            rows=rows,
            columns=columns,
            n_per_team=n_per_team,
            resize_image=resize_image,
            logs=logs,
        )
    except Exception as exc:
        return BlockFitResult(
            ordinal=block.ordinal,
            output_path=output_path,
            question=question,
            rows=None,
            columns=None,
            n_per_team=None,
            resize_image=None,
            logs=logs,
            error=str(exc),
        )


def print_block_result(result: BlockFitResult) -> None:
    print(f"{result.output_path.name}: fitting question {result.question}")
    for line in result.logs:
        print(line)
    if result.error:
        print(f"  ERROR: {result.error}", file=sys.stderr)
        return
    assert result.rows is not None
    assert result.columns is not None
    assert result.n_per_team is not None
    total = result.columns * result.rows
    resize = ""
    if result.resize_image is not None:
        resize = f", resize_image={format_float(result.resize_image)}"
    print(
        f"  final rows={result.rows}, total handouts={total}, "
        f"teams/page={total // result.n_per_team}{resize}"
    )


def updates_for_resize(
    block: HandoutBlock, resize_image: float | None
) -> dict[str, str | None]:
    resize_update = resize_update_for_block(block, resize_image)
    if resize_update is None:
        return {}
    return {"resize_image": format_float(resize_update)}


def update_source_resizes(
    source: Path,
    source_contents: str,
    blocks: list[HandoutBlock],
    resize_by_ordinal: dict[int, float | None],
) -> bool:
    updated_blocks = []
    changed = False
    for block in blocks:
        updates = updates_for_resize(block, resize_by_ordinal.get(block.ordinal))
        if updates:
            changed = True
            updated_blocks.append(upsert_metadata(block, updates).rstrip())
        else:
            updated_blocks.append(block.text.rstrip())
    if not changed:
        return False
    updated_contents = "\n---\n".join(updated_blocks) + "\n"
    if updated_contents == source_contents:
        return False
    source.write_text(updated_contents, encoding="utf8")
    return True


def all_q_block_text(
    block: HandoutBlock,
    source_dir: Path,
    output_dir: Path,
    resize_image: float | None,
) -> str:
    columns = parse_positive_int(block.meta.get("columns"), "columns")
    handouts_per_team = parse_positive_int(
        block.meta.get("handouts_per_team"), "handouts_per_team", default=3
    )
    updates = {"rows": str(valid_row_step(columns, handouts_per_team))}
    updates.update(updates_for_resize(block, resize_image))
    image_update = source_relative_image_update(block, source_dir, output_dir)
    if image_update is not None:
        updates["image"] = image_update
    return upsert_metadata(block, updates).rstrip()


def create_all_q_pdf(
    source: Path,
    output_dir: Path,
    blocks: list[HandoutBlock],
    resize_by_ordinal: dict[int, float | None],
    args,
) -> Path:
    output_pdf = output_dir / f"{source.stem}_all_q_1team.pdf"
    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(
            "w",
            encoding="utf8",
            suffix=".hndt",
            prefix=f".{source.stem}_all_q_1team_",
            dir=output_dir,
            delete=False,
        ) as tmp:
            temp_path = Path(tmp.name)
            tmp.write(
                "\n---\n".join(
                    all_q_block_text(
                        block,
                        source.parent,
                        output_dir,
                        resize_by_ordinal.get(block.ordinal),
                    )
                    for block in blocks
                )
                + "\n"
            )
        result = run_hndt2pdf(
            temp_path,
            args,
            output_pdf_path=output_pdf,
            compress_output=getattr(args, "compress_pdf", "off") == "on",
        )
        if not result.ok:
            details = result.stderr or result.stdout or "no compiler output"
            raise RuntimeError(f"all-q PDF compile failed:\n{details}")
        return output_pdf
    finally:
        if temp_path is not None:
            try:
                temp_path.unlink()
            except OSError:
                pass


def enabled_option(value) -> bool:
    if isinstance(value, str):
        return value == "on"
    return bool(value)


def run_split_fit(args) -> int:
    if args.jobs < 1:
        print("--jobs must be >= 1", file=sys.stderr)
        return 1
    if args.image_bottom_space_row_ratio < 0:
        print("--image-bottom-space-row-ratio must be >= 0", file=sys.stderr)
        return 1
    if args.image_shrink_percent <= 0 or args.image_shrink_percent >= 100:
        print("--image-shrink-percent must be > 0 and < 100", file=sys.stderr)
        return 1
    if args.min_resize_image <= 0:
        print("--min-resize-image must be > 0", file=sys.stderr)
        return 1
    source = Path(args.filename).expanduser().resolve()
    output_dir = Path(args.output_dir).expanduser().resolve() if args.output_dir else source.parent
    output_dir.mkdir(parents=True, exist_ok=True)

    source_contents = source.read_text(encoding="utf8")
    blocks = parse_blocks(source_contents)
    if not blocks:
        print(f"No handouts found in {source}", file=sys.stderr)
        return 1
    resize_config = ImageResizeConfig(
        enabled=not args.no_auto_resize_images,
        bottom_space_row_ratio=args.image_bottom_space_row_ratio,
        shrink_percent=args.image_shrink_percent,
        min_resize_image=args.min_resize_image,
    )

    used_names: set[str] = set()
    errors: list[str] = []
    results: list[BlockFitResult] = []
    jobs = min(args.jobs, len(blocks))
    print(f"Fitting {len(blocks)} handout(s) with {jobs} job(s).")

    block_jobs = [
        (block, output_path_for_block(block, source.stem, output_dir, used_names))
        for block in blocks
    ]
    with ThreadPoolExecutor(max_workers=jobs) as executor:
        futures = [
            executor.submit(
                fit_block,
                block,
                output_path,
                source.parent,
                args,
                args.max_rows,
                resize_config,
                enabled_option(args.keep_pdfs),
                args.verbose,
            )
            for block, output_path in block_jobs
        ]
        for future in as_completed(futures):
            result = future.result()
            results.append(result)
            print_block_result(result)
            if result.error:
                errors.append(f"{result.output_path.name}: {result.error}")

    if errors:
        print("\nErrors:", file=sys.stderr)
        for error in errors:
            print(f"- {error}", file=sys.stderr)
        return 1

    resize_by_ordinal = {result.ordinal: result.resize_image for result in results}
    if not args.no_update_source_resize:
        try:
            if update_source_resizes(source, source_contents, blocks, resize_by_ordinal):
                print(f"Updated resize_image values in {source}")
        except Exception as exc:
            print(f"Could not update source resize_image values: {exc}", file=sys.stderr)
            return 1

    if not args.no_all_q_pdf:
        try:
            all_q_pdf = create_all_q_pdf(
                source=source,
                output_dir=output_dir,
                blocks=blocks,
                resize_by_ordinal=resize_by_ordinal,
                args=args,
            )
            print(f"All-questions PDF: {all_q_pdf}")
        except Exception as exc:
            print(f"Could not create all-questions PDF: {exc}", file=sys.stderr)
            return 1
    return 0
