from dataclasses import dataclass
from pathlib import Path
import math

import fitz


MM_TO_PT = 72.0 / 25.4

A4_WIDTH_PT = 210 * MM_TO_PT
A4_HEIGHT_PT = 297 * MM_TO_PT

A3_WIDTH_PT = 297 * MM_TO_PT
A3_HEIGHT_PT = 420 * MM_TO_PT

LETTER_WIDTH_PT = 8.5 * 72
LETTER_HEIGHT_PT = 11 * 72


@dataclass
class SelectionRect:
    x: float
    y: float
    width: float
    height: float


@dataclass
class ExportOptions:
    output_mode: str
    split_layout: str
    export_type: str = "a3_split"
    selection_fit_mode: str = "fit_width"
    page_orientation: str = "auto"
    safe_margin_mm: float = 5.0
    overlap_mm: float = 5.0
    paper_size: str = "a4"
    poster_scale: float = 1.0
    tile_cols: int | None = None
    tile_rows: int | None = None
    page_rotations: dict[int, int] | None = None
    add_numbers: bool = True
    add_cut_guides: bool = True
    quality_scale: float = 2.5


class PosterExporter:
    def export(
        self,
        source_pdf_path: str,
        output_pdf_path: str,
        selections: dict[int, SelectionRect],
        options: ExportOptions,
    ) -> None:
        source_path = Path(source_pdf_path)
        output_path = Path(output_pdf_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        src_doc = fitz.open(str(source_path))
        out_doc = fitz.open()

        try:
            if options.export_type == "selection_only":
                self._export_selection_only(src_doc, out_doc, selections, options)
            else:
                self._export_a3_split(src_doc, out_doc, selections, options)

            if len(out_doc) == 0:
                raise RuntimeError("No output pages were generated.")

            out_doc.save(str(output_path), garbage=4, deflate=True)

        finally:
            out_doc.close()
            src_doc.close()

    def _export_a3_split(
        self,
        src_doc: fitz.Document,
        out_doc: fitz.Document,
        selections: dict[int, SelectionRect],
        options: ExportOptions,
    ) -> None:
        # A4 + 100% is the classic Poster Print base: A3-sized output split into 2 A4 sheets.
        # Other paper sizes or scales use the scalable tiling engine.
        paper_size = str(getattr(options, "paper_size", "a4") or "a4").lower()
        poster_scale = float(getattr(options, "poster_scale", 1.0) or 1.0)

        # A3 from A4 preset is 141% linear scale and should still be 2 A4 sheets.
        # Use the classic 2-sheet split for A4 up to A3 scale.
        if not (paper_size == "a4" and poster_scale <= 1.42):
            self._export_poster_tiling(src_doc, out_doc, selections, options)
            return

        sheet_total = len(selections) * 2
        sheet_no = 1

        for page_index in sorted(selections.keys()):
            if page_index < 0 or page_index >= len(src_doc):
                continue

            src_page = src_doc.load_page(page_index)
            selection = selections[page_index]
            full_clip = self._selection_to_rect(selection)

            if full_clip.width <= 0 or full_clip.height <= 0:
                continue

            tile_clips, tile_names = self._make_tile_clips(full_clip, options)

            for tile_index, tile_clip in enumerate(tile_clips):
                page = out_doc.new_page(width=A4_WIDTH_PT, height=A4_HEIGHT_PT)

                target_rect = self._get_target_rect(
                    options=options,
                    page_width=A4_WIDTH_PT,
                    page_height=A4_HEIGHT_PT,
                )

                # Keep PDF text/vector graphics sharp by placing the original PDF region
                # directly instead of rasterizing it to an image.
                page.show_pdf_page(
                    target_rect,
                    src_doc,
                    page_index,
                    clip=tile_clip,
                    rotate=self._get_page_rotation(page_index, options),
                    keep_proportion=False,
                )

                if options.add_cut_guides:
                    self._draw_cut_guides(
                        page=page,
                        options=options,
                        tile_name=tile_names[tile_index],
                        page_width=A4_WIDTH_PT,
                        page_height=A4_HEIGHT_PT,
                    )

                if options.add_numbers:
                    self._draw_numbering(
                        page=page,
                        page_width=A4_WIDTH_PT,
                        page_height=A4_HEIGHT_PT,
                        sheet_no=sheet_no,
                        sheet_total=sheet_total,
                        source_page=page_index + 1,
                        tile_no=tile_index + 1,
                        tile_total=2,
                        tile_name=tile_names[tile_index],
                    )

                sheet_no += 1

    def _export_selection_only(
        self,
        src_doc: fitz.Document,
        out_doc: fitz.Document,
        selections: dict[int, SelectionRect],
        options: ExportOptions,
    ) -> None:
        sheet_total = len(selections)
        sheet_no = 1

        for page_index in sorted(selections.keys()):
            if page_index < 0 or page_index >= len(src_doc):
                continue

            selection = selections[page_index]
            clip = self._selection_to_rect(selection)

            if clip.width <= 0 or clip.height <= 0:
                continue

            page_width, page_height = self._get_selection_page_size(clip, options)
            page = out_doc.new_page(width=page_width, height=page_height)

            printable_rect = self._get_target_rect(
                options=options,
                page_width=page_width,
                page_height=page_height,
            )

            source_clip, target_rect = self._calculate_selection_mapping(
                clip=clip,
                printable_rect=printable_rect,
                options=options,
            )

            page.show_pdf_page(
                target_rect,
                src_doc,
                page_index,
                clip=source_clip,
                rotate=self._get_page_rotation(page_index, options),
                keep_proportion=True,
            )

            # Selection Only Mode intentionally outputs clean pages.
            # Sheet numbers and cut guide lines are ignored in this mode,
            # even if the UI checkboxes are enabled.

            sheet_no += 1

    def _get_page_rotation(self, page_index: int, options: ExportOptions) -> int:
        try:
            rotations = getattr(options, "page_rotations", None) or {}
            preview_rotation = int(rotations.get(page_index, 0)) % 360

            # PyQt preview rotation and PyMuPDF show_pdf_page rotation use opposite visual direction.
            # Invert the angle so exported PDF matches the preview.
            return (-preview_rotation) % 360

        except Exception:
            return 0


    def _selection_to_rect(self, selection: SelectionRect) -> fitz.Rect:
        return fitz.Rect(
            selection.x,
            selection.y,
            selection.x + selection.width,
            selection.y + selection.height,
        )

    def _get_selection_page_size(
        self,
        clip: fitz.Rect,
        options: ExportOptions,
    ) -> tuple[float, float]:
        orientation = options.page_orientation

        if orientation == "auto":
            if clip.width >= clip.height:
                orientation = "landscape"
            else:
                orientation = "portrait"

        if orientation == "landscape":
            return A4_HEIGHT_PT, A4_WIDTH_PT

        return A4_WIDTH_PT, A4_HEIGHT_PT

    def _calculate_selection_mapping(
        self,
        clip: fitz.Rect,
        printable_rect: fitz.Rect,
        options: ExportOptions,
    ) -> tuple[fitz.Rect, fitz.Rect]:
        mode = options.selection_fit_mode

        if mode == "fit_whole":
            return clip, self._fit_rect_keep_aspect(clip, printable_rect)

        if mode == "fit_width":
            target = fitz.Rect(
                printable_rect.x0,
                printable_rect.y0,
                printable_rect.x1,
                printable_rect.y0 + printable_rect.width * (clip.height / clip.width),
            )
            target = self._center_rect_vertically(target, printable_rect)
            return clip, target

        if mode == "fit_height":
            target = fitz.Rect(
                printable_rect.x0,
                printable_rect.y0,
                printable_rect.x0 + printable_rect.height * (clip.width / clip.height),
                printable_rect.y1,
            )
            target = self._center_rect_horizontally(target, printable_rect)
            return clip, target

        if mode == "fill_crop":
            cropped_clip = self._crop_source_to_target_aspect(clip, printable_rect)
            return cropped_clip, printable_rect

        return clip, self._fit_rect_keep_aspect(clip, printable_rect)

    def _fit_rect_keep_aspect(
        self,
        source_rect: fitz.Rect,
        target_rect: fitz.Rect,
    ) -> fitz.Rect:
        source_ratio = source_rect.width / source_rect.height
        target_ratio = target_rect.width / target_rect.height

        if source_ratio >= target_ratio:
            width = target_rect.width
            height = width / source_ratio
        else:
            height = target_rect.height
            width = height * source_ratio

        x0 = target_rect.x0 + (target_rect.width - width) / 2.0
        y0 = target_rect.y0 + (target_rect.height - height) / 2.0

        return fitz.Rect(x0, y0, x0 + width, y0 + height)

    def _center_rect_vertically(
        self,
        rect: fitz.Rect,
        bounds: fitz.Rect,
    ) -> fitz.Rect:
        height = rect.height

        if height > bounds.height:
            return bounds

        y0 = bounds.y0 + (bounds.height - height) / 2.0
        return fitz.Rect(rect.x0, y0, rect.x1, y0 + height)

    def _center_rect_horizontally(
        self,
        rect: fitz.Rect,
        bounds: fitz.Rect,
    ) -> fitz.Rect:
        width = rect.width

        if width > bounds.width:
            return bounds

        x0 = bounds.x0 + (bounds.width - width) / 2.0
        return fitz.Rect(x0, rect.y0, x0 + width, rect.y1)

    def _crop_source_to_target_aspect(
        self,
        source_rect: fitz.Rect,
        target_rect: fitz.Rect,
    ) -> fitz.Rect:
        source_ratio = source_rect.width / source_rect.height
        target_ratio = target_rect.width / target_rect.height

        if source_ratio > target_ratio:
            new_width = source_rect.height * target_ratio
            x0 = source_rect.x0 + (source_rect.width - new_width) / 2.0
            return fitz.Rect(
                x0,
                source_rect.y0,
                x0 + new_width,
                source_rect.y1,
            )

        new_height = source_rect.width / target_ratio
        y0 = source_rect.y0 + (source_rect.height - new_height) / 2.0
        return fitz.Rect(
            source_rect.x0,
            y0,
            source_rect.x1,
            y0 + new_height,
        )

    def _make_tile_clips(
        self,
        full_clip: fitz.Rect,
        options: ExportOptions,
    ) -> tuple[list[fitz.Rect], list[str]]:
        overlap_pt = 0.0

        if options.output_mode == "overlap":
            overlap_pt = options.overlap_mm * MM_TO_PT

        if options.split_layout == "top_bottom":
            mid_y = (full_clip.y0 + full_clip.y1) / 2.0

            top = fitz.Rect(
                full_clip.x0,
                full_clip.y0,
                full_clip.x1,
                min(full_clip.y1, mid_y + overlap_pt),
            )

            bottom = fitz.Rect(
                full_clip.x0,
                max(full_clip.y0, mid_y - overlap_pt),
                full_clip.x1,
                full_clip.y1,
            )

            return [top, bottom], ["TOP", "BOTTOM"]

        mid_x = (full_clip.x0 + full_clip.x1) / 2.0

        left = fitz.Rect(
            full_clip.x0,
            full_clip.y0,
            min(full_clip.x1, mid_x + overlap_pt),
            full_clip.y1,
        )

        right = fitz.Rect(
            max(full_clip.x0, mid_x - overlap_pt),
            full_clip.y0,
            full_clip.x1,
            full_clip.y1,
        )

        return [left, right], ["LEFT", "RIGHT"]

    def _render_tile_to_a4(
        self,
        src_page: fitz.Page,
        full_clip: fitz.Rect,
        tile_clip: fitz.Rect,
        options: ExportOptions,
    ) -> fitz.Pixmap:
        zoom = self._calculate_a3_split_zoom(full_clip, options)
        high_quality_zoom = zoom * max(options.quality_scale, 1.0)
        matrix = fitz.Matrix(high_quality_zoom, high_quality_zoom)

        return src_page.get_pixmap(
            matrix=matrix,
            clip=tile_clip,
            alpha=False,
            colorspace=fitz.csRGB,
        )

    def _calculate_a3_split_zoom(
        self,
        full_clip: fitz.Rect,
        options: ExportOptions,
    ) -> float:
        if full_clip.width <= 0 or full_clip.height <= 0:
            raise RuntimeError("Invalid selection area.")

        if options.split_layout == "top_bottom":
            target_w_pt = A4_WIDTH_PT
            target_h_pt = A3_HEIGHT_PT
        else:
            target_w_pt = A3_WIDTH_PT
            target_h_pt = A4_HEIGHT_PT

        if options.output_mode == "no_cut":
            scale_factor = 0.96
        else:
            scale_factor = 1.0

        zoom_x = target_w_pt / full_clip.width
        zoom_y = target_h_pt / full_clip.height

        return min(zoom_x, zoom_y) * scale_factor

    def _get_target_rect(
        self,
        options: ExportOptions,
        page_width: float,
        page_height: float,
    ) -> fitz.Rect:
        if options.output_mode == "true_size":
            margin = 0
        elif options.output_mode == "overlap":
            margin = 3 * MM_TO_PT
        else:
            margin = options.safe_margin_mm * MM_TO_PT

        return fitz.Rect(
            margin,
            margin,
            page_width - margin,
            page_height - margin,
        )


    def _get_output_paper_size(self, options: ExportOptions) -> tuple[float, float]:
        paper_size = str(getattr(options, "paper_size", "a4") or "a4").lower()

        if paper_size in ("letter", "us_letter", "us letter"):
            return LETTER_WIDTH_PT, LETTER_HEIGHT_PT

        return A4_WIDTH_PT, A4_HEIGHT_PT

    def _get_poster_tiling_printable_rect(
        self,
        options: ExportOptions,
        page_width: float,
        page_height: float,
    ) -> fitz.Rect:
        margin = max(float(getattr(options, "safe_margin_mm", 5.0) or 0.0), 0.0) * MM_TO_PT

        # Keep a small minimum margin so cut guides and numbers remain visible.
        margin = max(margin, 3 * MM_TO_PT)

        return fitz.Rect(
            margin,
            margin,
            page_width - margin,
            page_height - margin,
        )

    def _make_poster_tiles(
        self,
        full_clip: fitz.Rect,
        options: ExportOptions,
        printable_rect: fitz.Rect,
    ) -> list[dict]:
        scale = max(float(getattr(options, "poster_scale", 1.0) or 1.0), 0.1)
        overlap_pt = 0.0

        if getattr(options, "output_mode", "") == "overlap":
            overlap_pt = max(float(getattr(options, "overlap_mm", 0.0) or 0.0), 0.0) * MM_TO_PT

        forced_cols = getattr(options, "tile_cols", None)
        forced_rows = getattr(options, "tile_rows", None)

        tiles: list[dict] = []

        # Preset-driven fixed grid, e.g. A2 from A4 = 2 x 2 = 4 sheets.
        if forced_cols and forced_rows and int(forced_cols) > 0 and int(forced_rows) > 0:
            cols = int(forced_cols)
            rows = int(forced_rows)

            tile_w = full_clip.width / cols
            tile_h = full_clip.height / rows

            overlap_source = overlap_pt / scale if overlap_pt > 0 else 0.0

            for row in range(rows):
                for col in range(cols):
                    x0 = full_clip.x0 + col * tile_w
                    y0 = full_clip.y0 + row * tile_h
                    x1 = full_clip.x0 + (col + 1) * tile_w
                    y1 = full_clip.y0 + (row + 1) * tile_h

                    # Add overlap toward neighboring tiles.
                    clip = fitz.Rect(
                        max(full_clip.x0, x0 - (overlap_source if col > 0 else 0)),
                        max(full_clip.y0, y0 - (overlap_source if row > 0 else 0)),
                        min(full_clip.x1, x1 + (overlap_source if col < cols - 1 else 0)),
                        min(full_clip.y1, y1 + (overlap_source if row < rows - 1 else 0)),
                    )

                    target_rect = printable_rect

                    tiles.append(
                        {
                            "row": row,
                            "col": col,
                            "rows": rows,
                            "cols": cols,
                            "clip": clip,
                            "target_rect": target_rect,
                        }
                    )

            return tiles

        printable_w = max(printable_rect.width, 1.0)
        printable_h = max(printable_rect.height, 1.0)

        # In output coordinates, each sheet covers printable_w/h.
        # Convert back to source PDF coordinates by dividing by scale.
        source_tile_w = printable_w / scale
        source_tile_h = printable_h / scale

        source_step_w = max((printable_w - overlap_pt) / scale, source_tile_w * 0.25)
        source_step_h = max((printable_h - overlap_pt) / scale, source_tile_h * 0.25)

        cols = max(1, math.ceil(max(full_clip.width - source_tile_w, 0.0) / source_step_w) + 1)
        rows = max(1, math.ceil(max(full_clip.height - source_tile_h, 0.0) / source_step_h) + 1)

        for row in range(rows):
            for col in range(cols):
                x0 = full_clip.x0 + col * source_step_w
                y0 = full_clip.y0 + row * source_step_h
                x1 = min(x0 + source_tile_w, full_clip.x1)
                y1 = min(y0 + source_tile_h, full_clip.y1)

                if x1 <= x0 or y1 <= y0:
                    continue

                clip = fitz.Rect(x0, y0, x1, y1)

                # Target size reflects true scale. Last row/column may not fill whole printable area.
                target_w = clip.width * scale
                target_h = clip.height * scale

                target_rect = fitz.Rect(
                    printable_rect.x0,
                    printable_rect.y0,
                    min(printable_rect.x0 + target_w, printable_rect.x1),
                    min(printable_rect.y0 + target_h, printable_rect.y1),
                )

                tiles.append(
                    {
                        "row": row,
                        "col": col,
                        "rows": rows,
                        "cols": cols,
                        "clip": clip,
                        "target_rect": target_rect,
                    }
                )

        return tiles

    def _export_poster_tiling(
        self,
        src_doc: fitz.Document,
        out_doc: fitz.Document,
        selections: dict[int, SelectionRect],
        options: ExportOptions,
    ) -> None:
        paper_width, paper_height = self._get_output_paper_size(options)
        printable_rect = self._get_poster_tiling_printable_rect(
            options=options,
            page_width=paper_width,
            page_height=paper_height,
        )

        all_tiles: list[tuple[int, fitz.Page, dict]] = []

        for page_index in sorted(selections.keys()):
            if page_index < 0 or page_index >= len(src_doc):
                continue

            src_page = src_doc.load_page(page_index)
            full_clip = self._selection_to_rect(selections[page_index])

            if full_clip.width <= 0 or full_clip.height <= 0:
                continue

            tiles = self._make_poster_tiles(
                full_clip=full_clip,
                options=options,
                printable_rect=printable_rect,
            )

            for tile in tiles:
                all_tiles.append((page_index, src_page, tile))

        sheet_total = len(all_tiles)
        sheet_no = 1

        for page_index, src_page, tile in all_tiles:
            page = out_doc.new_page(width=paper_width, height=paper_height)

            page.show_pdf_page(
                tile["target_rect"],
                src_doc,
                page_index,
                clip=tile["clip"],
                rotate=self._get_page_rotation(page_index, options),
                keep_proportion=False,
            )

            if options.add_cut_guides:
                self._draw_poster_tiling_cut_guides(
                    page=page,
                    page_width=paper_width,
                    page_height=paper_height,
                    printable_rect=printable_rect,
                    tile=tile,
                    options=options,
                )

            if options.add_numbers:
                self._draw_poster_tiling_numbering(
                    page=page,
                    page_width=paper_width,
                    page_height=paper_height,
                    sheet_no=sheet_no,
                    sheet_total=sheet_total,
                    source_page=page_index + 1,
                    tile=tile,
                )

            sheet_no += 1

    def _draw_poster_tiling_cut_guides(
        self,
        page: fitz.Page,
        page_width: float,
        page_height: float,
        printable_rect: fitz.Rect,
        tile: dict,
        options: ExportOptions,
    ) -> None:
        color = (0.85, 0.0, 0.0)
        width = 0.6
        dash = "[3 3] 0"

        row = int(tile.get("row", 0))
        col = int(tile.get("col", 0))
        rows = int(tile.get("rows", 1))
        cols = int(tile.get("cols", 1))

        # Outer printable trim rectangle.
        page.draw_rect(
            printable_rect,
            color=color,
            width=width,
            dashes=dash,
        )

        label_margin = 5 * MM_TO_PT

        # Join guide lines only on sides that have neighboring sheets.
        if col > 0:
            x = printable_rect.x0
            page.draw_line(
                fitz.Point(x, printable_rect.y0),
                fitz.Point(x, printable_rect.y1),
                color=color,
                width=width,
                dashes=dash,
            )
            self._draw_cut_label(page, "JOIN LEFT", x + 3, label_margin, page_width)

        if col < cols - 1:
            x = printable_rect.x1
            page.draw_line(
                fitz.Point(x, printable_rect.y0),
                fitz.Point(x, printable_rect.y1),
                color=color,
                width=width,
                dashes=dash,
            )
            self._draw_cut_label(page, "JOIN RIGHT", x - 78, label_margin, page_width)

        if row > 0:
            y = printable_rect.y0
            page.draw_line(
                fitz.Point(printable_rect.x0, y),
                fitz.Point(printable_rect.x1, y),
                color=color,
                width=width,
                dashes=dash,
            )
            self._draw_cut_label(page, "JOIN TOP", label_margin, y + 2, page_width)

        if row < rows - 1:
            y = printable_rect.y1
            page.draw_line(
                fitz.Point(printable_rect.x0, y),
                fitz.Point(printable_rect.x1, y),
                color=color,
                width=width,
                dashes=dash,
            )
            self._draw_cut_label(page, "JOIN BOTTOM", label_margin, y - 10, page_width)

        # If overlap mode is active, draw an inner overlap reference line.
        if getattr(options, "output_mode", "") == "overlap":
            overlap_pt = max(float(getattr(options, "overlap_mm", 0.0) or 0.0), 0.0) * MM_TO_PT

            if overlap_pt > 0:
                guide_color = (1.0, 0.45, 0.0)

                if col > 0:
                    x = min(printable_rect.x0 + overlap_pt, printable_rect.x1)
                    page.draw_line(
                        fitz.Point(x, printable_rect.y0),
                        fitz.Point(x, printable_rect.y1),
                        color=guide_color,
                        width=0.5,
                        dashes="[2 2] 0",
                    )

                if row > 0:
                    y = min(printable_rect.y0 + overlap_pt, printable_rect.y1)
                    page.draw_line(
                        fitz.Point(printable_rect.x0, y),
                        fitz.Point(printable_rect.x1, y),
                        color=guide_color,
                        width=0.5,
                        dashes="[2 2] 0",
                    )

    def _draw_poster_tiling_numbering(
        self,
        page: fitz.Page,
        page_width: float,
        page_height: float,
        sheet_no: int,
        sheet_total: int,
        source_page: int,
        tile: dict,
    ) -> None:
        # Minimal and consistent sheet label.
        # Example: 1/12, 2/12, ...
        text = f"{sheet_no}/{sheet_total}"

        point = fitz.Point(
            8 * MM_TO_PT,
            10 * MM_TO_PT,
        )

        page.insert_text(
            point,
            text,
            fontsize=8,
            fontname="helv",
            color=(0.35, 0.35, 0.35),
        )




    def _draw_cut_guides(
        self,
        page: fitz.Page,
        options: ExportOptions,
        tile_name: str,
        page_width: float,
        page_height: float,
    ) -> None:
        color = (0.85, 0.0, 0.0)
        width = 0.6
        dash = "[3 3] 0"

        margin = 4 * MM_TO_PT
        label_margin = 6 * MM_TO_PT

        if options.split_layout == "top_bottom":
            if tile_name == "TOP":
                y = page_height - margin
                page.draw_line(
                    fitz.Point(margin, y),
                    fitz.Point(page_width - margin, y),
                    color=color,
                    width=width,
                    dashes=dash,
                )
                self._draw_cut_label(page, "CUT / JOIN LINE", label_margin, y - 10, page_width)
            else:
                y = margin
                page.draw_line(
                    fitz.Point(margin, y),
                    fitz.Point(page_width - margin, y),
                    color=color,
                    width=width,
                    dashes=dash,
                )
                self._draw_cut_label(page, "CUT / JOIN LINE", label_margin, y + 2, page_width)
            return

        if tile_name == "LEFT":
            x = page_width - margin
            page.draw_line(
                fitz.Point(x, margin),
                fitz.Point(x, page_height - margin),
                color=color,
                width=width,
                dashes=dash,
            )
            self._draw_cut_label(page, "CUT / JOIN LINE", x - 70, label_margin, page_width)
        else:
            x = margin
            page.draw_line(
                fitz.Point(x, margin),
                fitz.Point(x, page_height - margin),
                color=color,
                width=width,
                dashes=dash,
            )
            self._draw_cut_label(page, "CUT / JOIN LINE", x + 3, label_margin, page_width)

    def _draw_cut_label(
        self,
        page: fitz.Page,
        text: str,
        x: float,
        y: float,
        page_width: float,
    ) -> None:
        rect = fitz.Rect(
            x,
            y,
            min(x + 90, page_width - 2 * MM_TO_PT),
            y + 10,
        )
        page.insert_textbox(
            rect,
            text,
            fontsize=6,
            fontname="helv",
            color=(0.85, 0.0, 0.0),
            align=fitz.TEXT_ALIGN_LEFT,
        )

    def _draw_selection_border(
        self,
        page: fitz.Page,
        target_rect: fitz.Rect,
    ) -> None:
        page.draw_rect(
            target_rect,
            color=(0.65, 0.65, 0.65),
            width=0.4,
        )

    def _draw_numbering(
        self,
        page: fitz.Page,
        page_width: float,
        page_height: float,
        sheet_no: int,
        sheet_total: int,
        source_page: int,
        tile_no: int,
        tile_total: int,
        tile_name: str,
    ) -> None:
        # Minimal and consistent sheet label.
        # Keep the same position on every page so users can find it easily.
        text = f"{sheet_no}/{sheet_total}"

        point = fitz.Point(
            8 * MM_TO_PT,
            10 * MM_TO_PT,
        )

        page.insert_text(
            point,
            text,
            fontsize=8,
            fontname="helv",
            color=(0.35, 0.35, 0.35),
        )




    def _draw_selection_only_numbering(
        self,
        page: fitz.Page,
        page_width: float,
        page_height: float,
        sheet_no: int,
        sheet_total: int,
        source_page: int,
    ) -> None:
        text = (
            f"Selection Only | "
            f"Sheet {sheet_no}/{sheet_total} | "
            f"Source {source_page}"
        )

        rect = fitz.Rect(
            8 * MM_TO_PT,
            page_height - 12 * MM_TO_PT,
            page_width - 8 * MM_TO_PT,
            page_height - 5 * MM_TO_PT,
        )

        page.insert_textbox(
            rect,
            text,
            fontsize=7,
            fontname="helv",
            color=(0.35, 0.35, 0.35),
            align=fitz.TEXT_ALIGN_RIGHT,
        )