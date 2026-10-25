"""
Export of the currently-loaded presentation to the two per-day-mode PDFs
(and, optionally, PNG image(s) - see ``export_day_pngs``).

Why this exists: the signage deck is a single .pptx file whose slides are
tagged #CHOLONLY / #NONCHOLONLY / #HIDDEN (see presentation.py's
_reconcile_day_modes and _apply_day_mode_visibility). Per the deployment
site's workflow, the same deck has to be published as TWO static PDFs for
people who can't watch the live slideshow:

  * luah_weekday.pdf - the chol (weekday) deck: every slide marked
    #CHOLONLY plus every always-visible slide.
  * luah_shabbos.pdf - the non-chol (Shabbos/Yom-Tov) deck: every slide
    marked #NONCHOLONLY plus every always-visible slide.

Slides tagged #HIDDEN, or #CONFIG (which presentation.py auto-hides, and
therefore always lands in the DAY_MODE_ALWAYS_HIDDEN bucket), are never
included in either PDF (or PNG).

PNG export (optional, off by default - see LuahConfig.export_png) reuses
the exact same day-mode subset selection, but via a raster filter that can
only ever hold ONE slide per file - see ``export_day_pngs``'s docstring for
the "preview" vs "all_slides" output shape this implies.

"Current slides (with modified elements)": the values already rendered into
the LIVE in-memory Impress document (tagged text substituted, analog-clock
shapes built, day-mode visibility applied) - NOT the raw template text in
the source .pptx on disk (we never write rendered values back to it, see
presentation.py's module docstring). So this module exports from the
already-loaded UNO document object directly, reusing LibreOffice's own PDF
renderer so the output matches what's actually displayed.

Subset selection: HOWEVER, this module does NOT rely on the PDF filter's
``PageRange`` media-descriptor to pick the subset - empirical testing showed
LibreOffice's ``impress_pdf_Export`` IGNORES ``PageRange`` (and also IGNORES
the slides' ``Visible`` property) and renders the WHOLE deck to every PDF.
Instead, each day-mode PDF is produced from its OWN, dedicated subset
presentation: we build a temporary Impress document containing copies of
exactly the slides that belong in that PDF (the ``#CHOLONLY``/always-visible
slides for the weekday deck, the ``#NONCHOLONLY``/always-visible slides for
the shabbos deck), export THAT whole document to PDF, and discard it. The
source (live) document is left completely unmodified. Uses
presentation_pdf._build_subset_document (slide copies preserve the rendered
content because we copy the already-rendered in-memory shapes).
"""

from __future__ import annotations

from pathlib import Path

# The per-slide day-mode buckets a slide can land in - kept in sync with
# presentation.py's DAY_MODE_* constants (referenced by name here via the
# controller rather than hardcoded re-import to avoid an import cycle).
DAY_MODE_CHOL_ONLY = "CHOL_ONLY"
DAY_MODE_NONCHOL_ONLY = "NONCHOL_ONLY"
DAY_MODE_ALWAYS_HIDDEN = "ALWAYS_HIDDEN"


def slide_indices_for_day(
    slide_count: int,
    slide_day_mode: dict,
    non_chol: bool,
) -> list[int]:
    """Pure, testable: computes the 0-based slide indices that belong in one
    of the two per-day-mode PDFs.

    Args:
        slide_count: total number of slides in the deck.
        slide_day_mode: the presentation's per-slide day-mode dict (slide
            index -> DAY_MODE_CHOL_ONLY / DAY_MODE_NONCHOL_ONLY /
            DAY_MODE_ALWAYS_HIDDEN, exactly as produced by
            presentation._reconcile_day_modes). A slide NOT present in the
            dict is an always-visible slide.
        non_chol: if True, compute the non-chol (shabbos) PDF's slide list;
            if False, the chol (weekday) PDF's.

    Returns the 0-based slide indices, in ascending deck order, for the
    requested day PDF: every always-visible slide ($index in slide_day_mode)
    PLUS every slide tagged for the requested day (DAY_MODE_NONCHOL_ONLY
    when non_chol=True, DAY_MODE_CHOL_ONLY when non_chol=False). Slides
    tagged DAY_MODE_ALWAYS_HIDDEN (which includes #CONFIG slides - see
    presentation.py) are never included, regardless of day."""
    wanted = (
        DAY_MODE_NONCHOL_ONLY if non_chol else DAY_MODE_CHOL_ONLY
    )
    indices: list[int] = []
    for index in range(slide_count):
        mode = slide_day_mode.get(index)
        if mode is None:  # always-visible (and not hidden by any mechanism)
            indices.append(index)
        elif mode == wanted:
            indices.append(index)
        # else: DAY_MODE_ALWAYS_HIDDEN, or the OTHER day's only-mode -
        # excluded from this PDF.
    return indices


def to_page_range(indices: list[int]) -> str:
    """Converts ascending 0-based slide indices into the 1-based,
    comma-separated pages/ranges string LibreOffice's PDF-export
    `PageRange` media-descriptor property accepts (e.g. [0, 2, 3, 6] ->
    "1,3-4,7"). Contiguous runs are collapsed into "a-b" ranges; individual
    pages are listed alone. Empty input produces ""."""
    if not indices:
        return ""
    # 0-based -> 1-based for the PageRange string.
    pages = [i + 1 for i in sorted(indices)]
    parts: list[str] = []
    run_start = pages[0]
    prev = pages[0]
    for page in pages[1:]:
        if page == prev + 1:
            prev = page
            continue
        parts.append(str(run_start) if run_start == prev else f"{run_start}-{prev}")
        run_start = prev = page
    parts.append(str(run_start) if run_start == prev else f"{run_start}-{prev}")
    return ",".join(parts)


def resolve_output_dir(config, attr_name: str = "export_pdf_dir") -> Path:
    """Resolves the directory a day-mode export is written to:
    ``getattr(config, attr_name)`` (a path string, e.g. ``export_pdf_dir``
    or ``export_png_dir``) if set, else the repo root (the directory that
    holds config.json / launch.py). Imported lazily to avoid an import
    cycle at module load time."""
    raw = getattr(config, attr_name, None)
    if raw:
        return Path(raw)

    from . import config as config_module

    return config_module.REPO_ROOT


def export_day_pdfs(
    desktop,
    document,
    slides,
    slide_day_mode: dict,
    output_dir: Path,
    weekday_name: str = "luah_weekday.pdf",
    shabbos_name: str = "luah_shabbos.pdf",
    overwrite: bool = True,
) -> list[Path]:
    """Exports the CURRENTLY-LOADED UNO ``document`` to the two day-mode
    PDFs (see module docstring): ``luah_weekday.pdf`` gets the `#CHOLONLY`
    + always-visible slides, ``luah_shabbos.pdf`` the `#NONCHOLONLY` +
    always-visible slides.

    ``desktop`` is LibreOffice's ``XDesktop`` (used to spin up the temporary
    subset documents). ``slides`` is the ordered list of the source deck's
    UNO draw-slide objects (presentation.PresentationController.
    _slides_by_index). ``slide_day_mode`` is the per-slide day-mode dict
    from the controller (see ``slide_indices_for_day``).

    The subset is NOT expressed via the PDF filter's ``PageRange`` media
    descriptor - empirical testing showed LibreOffice's ``impress_pdf_Export``
    IGNORES ``PageRange`` (and the slides' ``Visible`` property) and renders
    the whole deck into every PDF. Instead each day-mode PDF is produced by
    building a temporary subset presentation (see ``_build_subset_document``)
    that contains exact COPIES of only the target slides (so the rendered
    values are preserved) and exporting THAT whole document to PDF; the
    source ``document`` is left completely unmodified.

    Writes ``<output_dir>/<weekday_name>`` and ``<output_dir>/<shabbos_name>``
    (names overridable for testing) and returns the two output paths. Each
    export overwrites any previous file of the same name. UNO-dependent (not
    unit-testable here - no LibreOffice), like the rest of the UNO-touching
    code in this project."""
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    weekday_indices = slide_indices_for_day(len(slides), slide_day_mode, non_chol=False)
    shabbos_indices = slide_indices_for_day(len(slides), slide_day_mode, non_chol=True)

    written: list[Path] = []
    for include_indices, filename in (
        (weekday_indices, weekday_name),
        (shabbos_indices, shabbos_name),
    ):
        out_path = output_dir / filename
        if overwrite and out_path.exists():
            out_path.unlink()
        subset = None
        try:
            subset = _build_subset_document(desktop, document, slides, set(include_indices))
            _store_to_pdf(subset, out_path, overwrite=False)
            written.append(out_path)
        except Exception as exc:
            # A failure producing one day's PDF must not prevent the OTHER
            # day's PDF from being produced - log and continue.
            print(f"luah_signage: {filename} export failed (skipped): {exc!r}")
        finally:
            if subset is not None:
                try:
                    subset.close(False)
                except Exception:
                    pass
    return written


def _build_subset_document(desktop, document, slides, include_indices: set):
    """Returns a TEMPORARY Impress presentation containing only the source
    slides whose 0-based index is in ``include_indices`` (in deck order), with
    all the rendered content preserved.

    This works by a snapshot->filter round trip that is far more robust than
    per-shape copying:
      1. the live, already-rendered in-memory ``document`` is saved to a
         throwaway ODP file (LibreOffice's native format preserves the
         substituted tag text, clock shapes, etc. exactly);
      2. that ODP is reopened as a fresh, independent document (so removing
         slides here can never touch the live ``document``);
      3. every slide NOT in ``include_indices`` is removed from the reopened
         copy via ``XDrawPages.remove``;
      4. every remaining slide is forced ``Visible`` (the live document's
         day-mode may have hidden them - e.g. the single #CHOLONLY slide on a
         non-chol day - which would otherwise make an all-hidden subset fail
         to export to PDF with a 0xc10 Io/Write error).
    The source ``document`` is left completely unmodified. The caller exports
    and closes the returned subset document (and the temp ODP files are
    deleted by the caller)."""
    import tempfile
    from pathlib import Path as _P

    tmp_odp = _P(tempfile.gettempdir()) / "luah_signage_subset.odp"
    if tmp_odp.exists():
        tmp_odp.unlink()

    # 1) Snapshot the live rendered document to ODP (captures rendered values).
    _store_to_odp(document, tmp_odp, overwrite=True)

    # 2) Reopen as an independent copy. (We deliberately do NOT delete the
    # temp ODP here - on Windows an open document holds a lock on the file,
    # so the file is left for the OS temp clean-up to remove.)
    url = __import__("uno").systemPathToFileUrl(str(tmp_odp.resolve()))
    subset = desktop.loadComponentFromURL(
        url, "_blank", 0,
        (_prop("Hidden", True), _prop("ReadOnly", False)),
    )

    # 3) Remove every slide that does NOT belong in this subset PDF.
    pages = subset.DrawPages
    for i in reversed(range(pages.Count)):
        if i not in include_indices:
            try:
                pages.remove(pages.getByIndex(i))
            except Exception:
                pass

    # 4) Force every remaining slide VISIBLE.
    #
    # This is essential: the live source document has day-mode visibility
    # applied to it, so on a given day the slides that belong in THIS subset
    # may currently be marked Visible=False in the live doc (e.g. on a non-chol
    # day the single #CHOLONLY weekday slide that goes into luah_weekday.pdf is
    # hidden). The ODP snapshot reopens with those same hidden flags intact, so
    # without this a subset can end up as 0 visible slides and LibreOffice's
    # impress_pdf_Export fails with a 0xc10 Io/Write error instead of producing
    # the (small but legitimate) subset PDF. The two PDFs are static "this day's
    # deck" publications, so every included slide must be shown.
    for i in range(pages.Count):
        try:
            pages.getByIndex(i).Visible = True
        except Exception:
            pass
    return subset


def _store_to_odp(document, out_path: Path, overwrite: bool) -> None:
    """Saves ``document`` to ``out_path`` as an ODP (Impress's native format).
    Used to snapshot the live rendered document before filtering a subset (see
    ``_build_subset_document``)."""
    if overwrite and out_path.exists():
        out_path.unlink()
    url = __import__("uno").systemPathToFileUrl(str(out_path.resolve()))
    props = (
        _prop("FilterName", "impress8"),
        _prop("Overwrite", True),
    )
    document.storeToURL(url, props)


def _prop(name, value):
    """Builds a UNO ``PropertyValue`` for a media-descriptor entry."""
    from com.sun.star.beans import PropertyValue

    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def _store_to_pdf(document, out_path: Path, overwrite: bool) -> None:
    """Exports the ENTIRE ``document`` (which, for this feature, is a
    temporary subset presentation already containing only the target slides -
    see ``export_day_pdfs``/``_build_subset_document``) to ``out_path`` as a
    PDF. Removes any existing file first when ``overwrite`` is True. Raises on
    export failures (the caller swallows/logs them - see
    PresentationController.export_pdfs)."""
    if overwrite and out_path.exists():
        out_path.unlink()

    url = __import__("uno").systemPathToFileUrl(str(out_path.resolve()))
    props = (
        _prop("FilterName", "impress_pdf_Export"),
        _prop("Overwrite", True),
    )
    document.storeToURL(url, props)


def export_day_pngs(
    desktop,
    document,
    slides,
    slide_day_mode: dict,
    output_dir: Path,
    mode: str = "preview",
    weekday_prefix: str = "luah_weekday",
    shabbos_prefix: str = "luah_shabbos",
    overwrite: bool = True,
) -> list[Path]:
    """Exports the CURRENTLY-LOADED UNO ``document`` to PNG image(s) per
    day-mode, reusing the exact same subset selection as ``export_day_pdfs``
    (see module docstring) - but via LibreOffice's raster
    ``impress_png_Export`` filter, which, unlike the PDF filter, only ever
    exports the document's CURRENT page, never a whole multi-slide deck in
    one call. Two ``mode``s (see LuahConfig.export_png_mode's docstring):

    - ``"preview"`` (default): writes ONE png per day
      (``<prefix>.png``) containing just the FIRST slide of that day's
      subset.
    - ``"all_slides"``: writes ONE numbered png PER SLIDE in that day's
      subset (``<prefix>_01.png``, ``<prefix>_02.png``, ...).

    Reuses ``_build_subset_document`` (the same rendered-content-preserving,
    forced-all-visible subset doc already used for the PDFs), so an exported
    image shows exactly what that slide looks like in the PDF. Returns the
    list of written paths."""
    if mode not in ("preview", "all_slides"):
        raise ValueError(f"export_png_mode must be 'preview' or 'all_slides', got {mode!r}")

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    weekday_indices = slide_indices_for_day(len(slides), slide_day_mode, non_chol=False)
    shabbos_indices = slide_indices_for_day(len(slides), slide_day_mode, non_chol=True)

    written: list[Path] = []
    for include_indices, prefix in (
        (weekday_indices, weekday_prefix),
        (shabbos_indices, shabbos_prefix),
    ):
        subset = None
        try:
            subset = _build_subset_document(desktop, document, slides, set(include_indices))
            pages = subset.DrawPages
            page_positions = [0] if mode == "preview" else list(range(pages.Count))
            for position in page_positions:
                if position >= pages.Count:
                    continue
                suffix = "" if mode == "preview" else f"_{position + 1:02d}"
                out_path = output_dir / f"{prefix}{suffix}.png"
                _store_page_to_png(subset, pages.getByIndex(position), out_path, overwrite=overwrite)
                written.append(out_path)
        except Exception as exc:
            print(f"luah_signage: {prefix} PNG export failed (skipped): {exc!r}")
        finally:
            if subset is not None:
                try:
                    subset.close(False)
                except Exception:
                    pass
    return written


def _store_page_to_png(document, page, out_path: Path, overwrite: bool) -> None:
    """Exports a single Impress slide (``page``) to ``out_path`` as a PNG.
    LibreOffice's ``impress_png_Export`` filter always exports only the
    document's CURRENT page - unlike ``_store_to_pdf``, there is no
    whole-document page-range concept for a raster filter - so ``page`` is
    selected via the controller immediately before exporting. Raises on
    export failures (the caller swallows/logs them - see
    ``export_day_pngs``/PresentationController.export_pngs)."""
    if overwrite and out_path.exists():
        out_path.unlink()
    document.CurrentController.setCurrentPage(page)

    url = __import__("uno").systemPathToFileUrl(str(out_path.resolve()))
    props = (
        _prop("FilterName", "impress_png_Export"),
        _prop("Overwrite", True),
    )
    document.storeToURL(url, props)