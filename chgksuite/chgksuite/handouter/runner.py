#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os
import subprocess
import sys
import tempfile
import time

import toml
from PIL import Image
from pypdf import PdfReader, PdfWriter
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer

from chgksuite.common import get_source_dirs, set_lastdir
from chgksuite.handouter.gen import generate_handouts
from chgksuite.handouter.pack import pack_handouts
from chgksuite.handouter.installer import get_tectonic_path, install_tectonic
from chgksuite.handouter.tex_internals import (
    EDGE_DASHED,
    EDGE_NONE,
    EDGE_SOLID,
    GREYTEXT,
    HEADER,
    IMG,
    IMGWIDTH,
    TIKZBOX_END,
    TIKZBOX_INNER,
    TIKZBOX_START,
)
from chgksuite.handouter.utils import (
    compress_pdf,
    optimize_raster_image_for_tex,
    parse_handouts,
    read_file,
    replace_ext,
    write_file,
)


def tex_image_path(image_path):
    return str(image_path).replace("\\", "/")


def rotate_image(image_path, direction):
    """Rotate an image or PDF 90 degrees and save to a temp file.
    direction: 'r' for right (clockwise), 'l' for left (counter-clockwise).
    Returns the path to the rotated temp file.
    """
    ext = os.path.splitext(image_path)[1].lower() or ".png"

    if ext == ".pdf":
        reader = PdfReader(image_path)
        writer = PdfWriter()
        angle = 270 if direction == "r" else 90
        for page in reader.pages:
            page.rotate(angle)
            writer.add_page(page)
        fd, tmp_path = tempfile.mkstemp(suffix=ext)
        os.close(fd)
        with open(tmp_path, "wb") as f:
            writer.write(f)
        return tmp_path

    img = Image.open(image_path)
    # PIL's rotate is counter-clockwise, so right = -90, left = 90
    angle = -90 if direction == "r" else 90
    rotated = img.rotate(angle, expand=True)
    fd, tmp_path = tempfile.mkstemp(suffix=ext)
    os.close(fd)
    rotated.save(tmp_path)
    return tmp_path


class HandoutGenerator:
    SPACE = 1.5  # mm
    DEFAULT_TIKZ_MM = 2  # mm

    def __init__(self, args):
        self.args = args
        self._temp_files = []
        self.optimize_images = getattr(args, "optimize_images", "on") == "on"
        filename = getattr(args, "filename", None)
        if not isinstance(filename, (str, bytes, os.PathLike)):
            filename = None
        self.input_dir = (
            os.path.dirname(os.path.abspath(filename)) if filename else os.getcwd()
        )
        _, resourcedir = get_source_dirs()
        self.labels = toml.loads(
            read_file(os.path.join(resourcedir, f"labels_{args.language}.toml"))
        )
        self.blocks = [self.get_header()]

    def get_header(self):
        header = HEADER
        header = (
            header.replace("<PAPERWIDTH>", str(self.args.paperwidth))
            .replace("<PAPERHEIGHT>", str(self.args.paperheight))
            .replace("<MARGIN_LEFT>", str(self.args.margin_left))
            .replace("<MARGIN_RIGHT>", str(self.args.margin_right))
            .replace("<MARGIN_TOP>", str(self.args.margin_top))
            .replace("<MARGIN_BOTTOM>", str(self.args.margin_bottom))
            .replace(
                "<TIKZ_MM>",
                str(
                    self.args.tikz_mm
                    if self.args.tikz_mm is not None
                    else self.DEFAULT_TIKZ_MM
                ),
            )
        )
        if self.args.font:
            header = header.replace("Arial", self.args.font)
        return header

    def parse_input(self, filepath):
        contents = read_file(filepath)
        return parse_handouts(contents)

    def generate_for_question(self, question_num):
        handout_text = self.labels["general"]["handout_for_question"].format(
            question_num
        )
        return GREYTEXT.replace("<GREYTEXT>", handout_text)

    def make_tikzbox(self, block, edges=None, ext=None, inner_sep=None):
        """
        Create a TikZ box with configurable edge styles and extensions.
        edges is a dict with keys 'top', 'bottom', 'left', 'right'
        values are EDGE_DASHED or EDGE_SOLID
        ext is a dict with edge extensions to close gaps at boundaries
        """
        if edges is None:
            edges = {
                "top": EDGE_DASHED,
                "bottom": EDGE_DASHED,
                "left": EDGE_DASHED,
                "right": EDGE_DASHED,
            }
        if ext is None:
            ext = {
                "top": ("0pt", "0pt"),
                "bottom": ("0pt", "0pt"),
                "left": ("0pt", "0pt"),
                "right": ("0pt", "0pt"),
            }

        if block.get("no_center"):
            align = ""
        else:
            align = ", align=center"
        textwidth = ", text width=\\boxwidthinner"
        fs = block.get("font_size") or self.args.font_size
        fontsize = "\\fontsize{FSpt}{LHpt}\\selectfont ".replace("FS", str(fs)).replace(
            "LH", str(round(fs * 1.2, 1))
        )
        contents = block["contents"]
        if block.get("font_family"):
            contents = "\\fontspec{" + block["font_family"] + "}" + contents
        inner_sep_str = f", inner sep={inner_sep}mm" if inner_sep is not None else ""
        return (
            TIKZBOX_INNER.replace("<CONTENTS>", contents)
            .replace("<ALIGN>", align)
            .replace("<TEXTWIDTH>", textwidth)
            .replace("<INNER_SEP_OVERRIDE>", inner_sep_str)
            .replace("<FONTSIZE>", fontsize)
            .replace("<TOP>", edges["top"])
            .replace("<BOTTOM>", edges["bottom"])
            .replace("<LEFT>", edges["left"])
            .replace("<RIGHT>", edges["right"])
            .replace("<TOP_EXT_L>", ext["top"][0])
            .replace("<TOP_EXT_R>", ext["top"][1])
            .replace("<BOTTOM_EXT_L>", ext["bottom"][0])
            .replace("<BOTTOM_EXT_R>", ext["bottom"][1])
            .replace("<LEFT_EXT_T>", ext["left"][0])
            .replace("<LEFT_EXT_B>", ext["left"][1])
            .replace("<RIGHT_EXT_T>", ext["right"][0])
            .replace("<RIGHT_EXT_B>", ext["right"][1])
        )

    def get_page_width(self):
        return self.args.paperwidth - self.args.margin_left - self.args.margin_right - 2

    def get_block_max_width(self, block):
        max_width = block.get("max_width", 1.0)
        if max_width <= 0 or max_width > 1:
            raise ValueError(f"max_width must be between 0 and 1, got {max_width}")
        return max_width

    def resolve_image_path(self, image_path):
        if os.path.isabs(image_path):
            return image_path
        return os.path.join(self.input_dir, image_path)

    def prepare_image(self, image_path):
        if not self.optimize_images:
            return image_path
        source_path = self.resolve_image_path(image_path)
        optimized_path = optimize_raster_image_for_tex(source_path, quality=80)
        if optimized_path != source_path:
            self._temp_files.append(optimized_path)
            return optimized_path
        return image_path

    def get_cut_direction(
        self, columns, num_rows, handouts_per_team, grouping="horizontal"
    ):
        """
        Determine team rectangle dimensions.
        Returns (team_cols, team_rows) where each team is a team_cols × team_rows block.

        Falls back to (None, None) if handouts can't be evenly divided into teams.

        Args:
            grouping: "horizontal" (default) prefers wider teams (smaller team_rows),
                      "vertical" prefers taller teams (smaller team_cols).
        """
        total = columns * num_rows

        # Check if total handouts can be evenly divided
        if total % handouts_per_team != 0:
            return None, None

        num_teams = total // handouts_per_team
        if num_teams < 1:
            return None, None  # Invalid configuration

        # Find all valid team rectangle sizes (team_cols × team_rows = handouts_per_team)
        valid_layouts = []
        for team_rows in range(1, handouts_per_team + 1):
            if handouts_per_team % team_rows == 0:
                team_cols = handouts_per_team // team_rows
                if columns % team_cols == 0 and num_rows % team_rows == 0:
                    valid_layouts.append((team_cols, team_rows))

        if not valid_layouts:
            return None, None

        # Sort based on grouping preference
        if grouping == "vertical":
            # Prefer vertical grouping (smaller team_cols = taller teams)
            valid_layouts.sort(key=lambda x: x[0])
        else:
            # Prefer horizontal grouping (smaller team_rows = wider teams)
            valid_layouts.sort(key=lambda x: x[1])

        return valid_layouts[0]

    def get_edge_styles(
        self,
        row_idx,
        col_idx,
        num_rows,
        columns,
        team_cols,
        team_rows,
        hspace=None,
        vspace=None,
    ):
        """
        Determine edge styles and extensions for a box at position (row_idx, col_idx).
        Outer edges of team rectangles are solid (thicker), inner edges are dashed.
        Extensions are used to close gaps in ALL solid lines.
        Duplicate dashed edges are skipped to avoid double lines.

        team_cols and team_rows define the dimensions of each team rectangle.
        """
        # Default: all dashed, no extension
        edges = {
            "top": EDGE_DASHED,
            "bottom": EDGE_DASHED,
            "left": EDGE_DASHED,
            "right": EDGE_DASHED,
        }
        ext = {
            "top": ("0pt", "0pt"),
            "bottom": ("0pt", "0pt"),
            "left": ("0pt", "0pt"),
            "right": ("0pt", "0pt"),
        }

        # Gap sizes (half of spacing to extend into)
        h_sp = hspace if hspace is not None else self.SPACE
        v_sp = vspace if vspace is not None else 1.0
        h_gap = f"{h_sp / 2}mm"
        v_gap = f"{v_sp / 2}mm"

        # Helper functions to check if position is at a team boundary
        def is_at_right_team_boundary():
            """Is this box at the right edge of its team (but not at grid edge)?"""
            if not team_cols:
                return False
            return (col_idx + 1) % team_cols == 0 and col_idx < columns - 1

        def is_at_left_team_boundary():
            """Is this box at the left edge of its team (but not at grid edge)?"""
            if not team_cols:
                return False
            return col_idx % team_cols == 0 and col_idx > 0

        def is_at_bottom_team_boundary():
            """Is this box at the bottom edge of its team (but not at grid edge)?"""
            if not team_rows:
                return False
            return (row_idx + 1) % team_rows == 0 and row_idx < num_rows - 1

        def is_at_top_team_boundary():
            """Is this box at the top edge of its team (but not at grid edge)?"""
            if not team_rows:
                return False
            return row_idx % team_rows == 0 and row_idx > 0

        # Determine which edges are solid
        # Only apply solid edges if we have valid team dimensions
        # Otherwise fall back to all-dashed (default)
        if team_cols is not None and team_rows is not None:
            # Outer edges of the entire grid
            if row_idx == 0:
                edges["top"] = EDGE_SOLID
            if row_idx == num_rows - 1:
                edges["bottom"] = EDGE_SOLID
            if col_idx == 0:
                edges["left"] = EDGE_SOLID
            if col_idx == columns - 1:
                edges["right"] = EDGE_SOLID

            # Team boundary edges
            if is_at_right_team_boundary():
                edges["right"] = EDGE_SOLID
            if is_at_left_team_boundary():
                edges["left"] = EDGE_SOLID
            if is_at_bottom_team_boundary():
                edges["bottom"] = EDGE_SOLID
            if is_at_top_team_boundary():
                edges["top"] = EDGE_SOLID

        # Skip duplicate dashed edges (to avoid double lines between adjacent boxes)
        if edges["left"] == EDGE_DASHED and col_idx > 0:
            edges["left"] = EDGE_NONE

        if edges["top"] == EDGE_DASHED and row_idx > 0:
            edges["top"] = EDGE_NONE

        # Calculate extensions for solid edges to close gaps
        # But don't extend into team boundary gaps!

        if edges["top"] == EDGE_SOLID:
            at_left_boundary = is_at_left_team_boundary()
            ext_left = "-" + h_gap if col_idx > 0 and not at_left_boundary else "0pt"
            at_right_boundary = is_at_right_team_boundary()
            ext_right = (
                h_gap if col_idx < columns - 1 and not at_right_boundary else "0pt"
            )
            ext["top"] = (ext_left, ext_right)

        if edges["bottom"] == EDGE_SOLID:
            at_left_boundary = is_at_left_team_boundary()
            ext_left = "-" + h_gap if col_idx > 0 and not at_left_boundary else "0pt"
            at_right_boundary = is_at_right_team_boundary()
            ext_right = (
                h_gap if col_idx < columns - 1 and not at_right_boundary else "0pt"
            )
            ext["bottom"] = (ext_left, ext_right)

        if edges["left"] == EDGE_SOLID:
            at_top_boundary = is_at_top_team_boundary()
            ext_top = v_gap if row_idx > 0 and not at_top_boundary else "0pt"
            at_bottom_boundary = is_at_bottom_team_boundary()
            ext_bottom = (
                "-" + v_gap
                if row_idx < num_rows - 1 and not at_bottom_boundary
                else "0pt"
            )
            ext["left"] = (ext_top, ext_bottom)

        if edges["right"] == EDGE_SOLID:
            at_top_boundary = is_at_top_team_boundary()
            ext_top = v_gap if row_idx > 0 and not at_top_boundary else "0pt"
            at_bottom_boundary = is_at_bottom_team_boundary()
            ext_bottom = (
                "-" + v_gap
                if row_idx < num_rows - 1 and not at_bottom_boundary
                else "0pt"
            )
            ext["right"] = (ext_top, ext_bottom)

        return edges, ext

    def generate_regular_block(self, block_):
        block = block_.copy()
        if not (block.get("image") or block.get("text")):
            return
        columns = block["columns"]
        num_rows = block.get("rows") or 1
        handouts_per_team = block.get("handouts_per_team") or 3
        grouping = block.get("grouping") or "horizontal"

        # Determine team rectangle dimensions
        team_cols, team_rows = self.get_cut_direction(
            columns, num_rows, handouts_per_team, grouping
        )
        if self.args.debug:
            print(
                f"team_cols: {team_cols}, team_rows: {team_rows}, grouping: {grouping}"
            )

        hspace = block.get("hspace") or self.SPACE
        vspace_val = block.get("vspace")
        tikz_mm = block.get("tikz_mm")

        spaces = columns - 1
        available_width = self.get_page_width() * self.get_block_max_width(block)
        boxwidth = self.args.boxwidth or round(
            (available_width - spaces * hspace) / columns,
            3,
        )
        total_width = boxwidth * columns + spaces * hspace
        if self.args.debug:
            print(
                f"columns: {columns}, boxwidth: {boxwidth}, total width: {total_width}"
            )
        if self.args.tikz_mm is not None:
            effective_tikz_mm = self.args.tikz_mm
        elif tikz_mm is not None:
            effective_tikz_mm = tikz_mm
        else:
            effective_tikz_mm = self.DEFAULT_TIKZ_MM
        boxwidthinner = self.args.boxwidthinner or (boxwidth - 2 * effective_tikz_mm)
        header = [
            r"\setlength{\boxwidth}{<Q>mm}%".replace("<Q>", str(boxwidth)),
            r"\setlength{\boxwidthinner}{<Q>mm}%".replace("<Q>", str(boxwidthinner)),
        ]
        contents = []
        if block.get("image"):
            image_path = block["image"]
            if block.get("rotate"):
                image_path = rotate_image(
                    self.resolve_image_path(image_path), block["rotate"]
                )
                self._temp_files.append(image_path)
            image_path = self.prepare_image(image_path)
            img_qwidth = block.get("resize_image") or 1.0
            imgwidth = IMGWIDTH.replace("<QWIDTH>", str(img_qwidth))
            contents.append(
                IMG.replace("<IMGPATH>", tex_image_path(image_path)).replace(
                    "<IMGWIDTH>", imgwidth
                )
            )
        if block.get("text"):
            contents.append(block["text"])
        block["contents"] = "\\linebreak\n\\hstrut ".join(contents)
        if block.get("no_center"):
            block["centering"] = ""
        else:
            block["centering"] = "\\centering"

        rows = []
        for row_idx in range(num_rows):
            row_boxes = []
            for col_idx in range(columns):
                edges, ext = self.get_edge_styles(
                    row_idx,
                    col_idx,
                    num_rows,
                    columns,
                    team_cols,
                    team_rows,
                    hspace=hspace,
                    vspace=vspace_val if vspace_val is not None else 1.0,
                )
                row_boxes.append(
                    self.make_tikzbox(block, edges, ext, inner_sep=effective_tikz_mm)
                )
            row = (
                TIKZBOX_START.replace("<CENTERING>", block["centering"])
                + "\n".join(row_boxes)
                + TIKZBOX_END
            )
            rows.append(row)
        vs = vspace_val if vspace_val is not None else 1
        return "\n".join(header) + "\n" + f"\n\n\\vspace{{{vs}mm}}\n\n".join(rows)

    def generate(self):
        for block in self.parse_input(self.args.filename):
            if not block:
                self.blocks.append("\n\\clearpage\n")
                continue
            if self.args.debug:
                print(block)
            if block.get("for_question"):
                self.blocks.append(self.generate_for_question(block["for_question"]))
            if block.get("columns"):
                block = self.generate_regular_block(block)
                if block:
                    self.blocks.append(block)
        self.blocks.append("\\end{document}")
        return "\n\n".join(self.blocks)


def get_num_teams(filepath):
    """Extract the number of teams from the first regular block of a .hndt file."""
    contents = read_file(filepath)
    blocks = parse_handouts(contents)
    for block in blocks:
        if block.get("columns"):
            columns = block["columns"]
            num_rows = block.get("rows") or 1
            handouts_per_team = block.get("handouts_per_team") or 3
            total = columns * num_rows
            if total % handouts_per_team == 0:
                return total // handouts_per_team
    return None


def process_file(args, file_dir, bn):
    generator = HandoutGenerator(args)
    tex_contents = generator.generate()
    add_n_teams = getattr(args, "add_n_teams", "off") == "on"
    num_teams = get_num_teams(args.filename) if add_n_teams else None
    if num_teams is not None:
        pdf_bn = f"{bn}_{num_teams}teams_{args.language}"
    else:
        pdf_bn = f"{bn}_{args.language}"
    tex_path = os.path.join(file_dir, f"{pdf_bn}.tex")
    write_file(tex_path, tex_contents)

    tectonic_path = get_tectonic_path()
    if not tectonic_path:
        print("tectonic is not present, installing it...")
        install_tectonic(args)
        tectonic_path = get_tectonic_path()
    if not tectonic_path:
        raise Exception("tectonic couldn't be installed successfully :(")
    if args.debug:
        print(f"tectonic found at `{tectonic_path}`")

    proc = subprocess.run(
        [tectonic_path, os.path.basename(tex_path)],
        check=False,
        cwd=file_dir,
        text=True,
        capture_output=True,
    )
    if proc.stdout:
        print(proc.stdout, end="")
    if proc.stderr:
        print(proc.stderr, end="", file=sys.stderr)
    proc.check_returncode()

    for tmp in generator._temp_files:
        try:
            os.remove(tmp)
        except OSError:
            pass

    output_file = replace_ext(tex_path, "pdf")

    if args.compress_pdf == "on":
        compress_pdf(output_file)

    print(f"Output file: {output_file}")

    if not args.debug:
        os.remove(tex_path)


class FileChangeHandler(FileSystemEventHandler):
    def __init__(self, args, file_dir, bn):
        self.args = args
        self.file_dir = file_dir
        self.bn = bn
        self.last_processed = 0

    def on_modified(self, event):
        if event.src_path == os.path.abspath(self.args.filename):
            # Debounce to avoid processing the same change multiple times
            current_time = time.time()
            if current_time - self.last_processed > 1:
                print(f"File {self.args.filename} changed, regenerating PDF...")
                process_file(self.args, self.file_dir, self.bn)
                self.last_processed = current_time


def run_handouter(args):
    file_dir = os.path.dirname(os.path.abspath(args.filename))
    bn, _ = os.path.splitext(os.path.basename(args.filename))

    process_file(args, file_dir, bn)

    if args.watch:
        print(f"Watching {args.filename} for changes. Press Ctrl+C to stop.")
        event_handler = FileChangeHandler(args, file_dir, bn)
        observer = Observer()
        observer.schedule(event_handler, path=file_dir, recursive=False)
        observer.start()
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            observer.stop()
        observer.join()


def gui_handouter(args):
    if hasattr(args, "filename") and args.filename:
        set_lastdir(os.path.dirname(os.path.abspath(args.filename)))
    if args.handoutssubcommand in ("4s2hndt", "generate"):
        generate_handouts(args)
    elif args.handoutssubcommand in ("hndt2pdf", "run"):
        run_handouter(args)
    elif args.handoutssubcommand == "install":
        install_tectonic(args)
    elif args.handoutssubcommand == "split_fit":
        from chgksuite.handouter.split_fit import run_split_fit

        exit_code = run_split_fit(args)
        if exit_code:
            raise SystemExit(exit_code)
    elif args.handoutssubcommand == "pack":
        pack_handouts(args)
    elif args.handoutssubcommand == "create_html":
        from chgksuite.handouter.html_handout import create_html

        create_html(args)
    elif args.handoutssubcommand == "html2img":
        from chgksuite.handouter.html_handout import html2img

        html2img(args)
