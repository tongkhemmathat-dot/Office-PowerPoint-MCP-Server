"""
Slide rendering tool for PowerPoint MCP Server.
Renders a slide to a PNG so the result can be inspected visually.

Renderers (tried in order for renderer="auto"):
  powerpoint   - Microsoft PowerPoint via COM (Windows only, PowerPoint must be installed)
  libreoffice  - `soffice --headless` (any OS, LibreOffice must be installed)
"""

import os
import shutil
import subprocess
import sys
import tempfile
from typing import Dict, List, Optional

from mcp.types import ToolAnnotations
from pptx import Presentation

_LIBREOFFICE_PATHS = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/bin/soffice",
    "/usr/local/bin/soffice",
]

_POWERPOINT_SCRIPT = r"""
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject PowerPoint.Application
try {
    $p = $app.Presentations.Open($env:PPT_RENDER_SRC, -1, 0, 0)
    try { $p.Slides.Item([int]$env:PPT_RENDER_SLIDE).Export($env:PPT_RENDER_OUT, 'PNG', [int]$env:PPT_RENDER_W, [int]$env:PPT_RENDER_H) }
    finally { $p.Close() }
} finally {
    # Leave PowerPoint running if the user has other presentations open.
    if ($app.Presentations.Count -eq 0) { $app.Quit() }
}
"""


def _find_soffice() -> Optional[str]:
    found = shutil.which("soffice") or shutil.which("libreoffice")
    if found:
        return found
    return next((p for p in _LIBREOFFICE_PATHS if os.path.exists(p)), None)


def _keep_only_slide(path: str, slide_index: int) -> None:
    """Rewrite the pptx at `path` so it contains only the given slide."""
    pres = Presentation(path)
    sld_id_lst = pres.slides._sldIdLst
    for i, sld_id in reversed(list(enumerate(list(sld_id_lst)))):
        if i != slide_index:
            pres.part.drop_rel(sld_id.rId)
            sld_id_lst.remove(sld_id)
    pres.save(path)


def _render_powerpoint(src: str, slide_index: int, out: str, width_px: int, height_px: int) -> None:
    if sys.platform != "win32":
        raise RuntimeError("PowerPoint renderer is only available on Windows")
    env = dict(os.environ, PPT_RENDER_SRC=src, PPT_RENDER_OUT=out,
               PPT_RENDER_SLIDE=str(slide_index + 1),
               PPT_RENDER_W=str(width_px), PPT_RENDER_H=str(height_px))
    proc = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", _POWERPOINT_SCRIPT],
        env=env, capture_output=True, text=True, timeout=120,
    )
    if proc.returncode != 0 or not os.path.exists(out):
        raise RuntimeError(f"PowerPoint export failed: {(proc.stderr or proc.stdout).strip()[:300]}")


def _render_libreoffice(src: str, slide_index: int, out: str) -> None:
    soffice = _find_soffice()
    if not soffice:
        raise RuntimeError("LibreOffice (soffice) not found")
    work = tempfile.mkdtemp(prefix="ppt_render_")
    try:
        single = os.path.join(work, "slide.pptx")
        shutil.copyfile(src, single)
        _keep_only_slide(single, slide_index)  # soffice's PNG export renders only the first slide
        profile = os.path.join(work, "profile")
        proc = subprocess.run(
            [soffice, f"-env:UserInstallation=file:///{profile.replace(os.sep, '/')}",
             "--headless", "--convert-to", "png", "--outdir", work, single],
            capture_output=True, text=True, timeout=180,
        )
        produced = os.path.join(work, "slide.png")
        if not os.path.exists(produced):
            raise RuntimeError(f"LibreOffice export failed: {(proc.stderr or proc.stdout).strip()[:300]}")
        shutil.move(produced, out)
    finally:
        shutil.rmtree(work, ignore_errors=True)


def register_render_tools(app, presentations, get_current_presentation_id):
    """Register rendering tools with the FastMCP app."""

    @app.tool(
        annotations=ToolAnnotations(
            title="Render Slide",
            readOnlyHint=True,
        ),
    )
    def render_slide(
        slide_index: int = 0,
        output_path: Optional[str] = None,
        width_px: int = 1600,
        renderer: str = "auto",
        return_image: bool = False,
        presentation_id: Optional[str] = None
    ):
        """
        Render a slide to a PNG image to check the layout visually.

        Needs Microsoft PowerPoint (Windows) or LibreOffice installed on the machine
        running the server. The presentation is rendered from its current in-memory
        state; nothing needs to be saved first.

        Args:
            slide_index: Index of the slide (0-based)
            output_path: Where to write the PNG (default: a temp file)
            width_px: Image width in pixels; height follows the slide aspect ratio
            renderer: "auto", "powerpoint" or "libreoffice"
            return_image: Also return the image itself (for clients that can display images,
                or when the server runs on another machine than the client)
            presentation_id: Optional presentation ID (uses current if not provided)
        """
        try:
            pres_id = presentation_id or get_current_presentation_id()
            if pres_id not in presentations:
                return {"error": "Presentation not found"}
            pres = presentations[pres_id]
            if not (0 <= slide_index < len(pres.slides)):
                return {"error": f"Slide index {slide_index} out of range"}
            if renderer not in ("auto", "powerpoint", "libreoffice"):
                return {"error": "renderer must be auto, powerpoint or libreoffice"}
            if not (100 <= width_px <= 8000):
                return {"error": "width_px must be between 100 and 8000"}

            height_px = round(width_px * int(pres.slide_height) / int(pres.slide_width))
            out = os.path.abspath(output_path) if output_path else os.path.join(
                tempfile.gettempdir(), f"{pres_id}_slide{slide_index}.png")
            os.makedirs(os.path.dirname(out), exist_ok=True)
            if os.path.exists(out):
                os.remove(out)

            work = tempfile.mkdtemp(prefix="ppt_src_")
            try:
                src = os.path.join(work, "render.pptx")
                pres.save(src)

                order: List[str] = ["powerpoint", "libreoffice"] if renderer == "auto" else [renderer]
                errors: List[str] = []
                used = None
                for name in order:
                    try:
                        if name == "powerpoint":
                            _render_powerpoint(src, slide_index, out, width_px, height_px)
                        else:
                            _render_libreoffice(src, slide_index, out)
                        used = name
                        break
                    except Exception as e:  # try the next renderer
                        errors.append(f"{name}: {e}")
            finally:
                shutil.rmtree(work, ignore_errors=True)

            if not used:
                return {"error": "No renderer available. " + " | ".join(errors)}

            result = {"image_path": out, "renderer": used, "slide_index": slide_index}
            if return_image:
                from mcp.server.fastmcp import Image
                return [result, Image(path=out)]
            return result

        except Exception as e:
            return {"error": f"Failed to render slide: {str(e)}"}
