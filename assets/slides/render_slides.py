"""Render every slide of a .pptx to PNG, so the slides can be looked at before anyone else sees them; on Windows
with PowerPoint, also check that PowerPoint lays out every text box exactly as built (no re-wrapped line, no
overflow).

    python render_slides.py deck.pptx out_folder [width_px]

Windows with PowerPoint: drives PowerPoint through pywin32. It opens a COPY of the deck read-only and without a
window, exports each slide, closes only that copy, and quits PowerPoint only if PowerPoint was not already running
(so a deck the user has open is never touched).
Otherwise (or if that fails): LibreOffice, if installed, with its own temporary profile, and PyMuPDF to turn the PDF
pages into PNGs. If neither is available, ask the user to export the slides from PowerPoint (File > Export > PNG).
"""
import os
import pathlib
import shutil
import subprocess
import sys
import tempfile


def powerpoint_running():
    try:
        out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE"], capture_output=True, text=True).stdout
        return "POWERPNT.EXE" in out
    except OSError:
        return False


def _with_powerpoint(deck, work):
    """Open a temporary copy of deck in PowerPoint (read-only, no window), run work(presentation), clean up."""
    import win32com.client as win32
    deck = pathlib.Path(deck).resolve()
    was_running = powerpoint_running()
    tmp = pathlib.Path(tempfile.mkdtemp()) / deck.name          # a copy, so the original is never locked
    shutil.copyfile(deck, tmp)
    app = win32.DispatchEx("PowerPoint.Application")
    pres = None
    try:
        pres = app.Presentations.Open(str(tmp), True, False, False)   # ReadOnly, Untitled=False, WithWindow=False
        return work(pres)
    finally:
        if pres is not None:
            pres.Close()
        try:
            if not was_running and app.Presentations.Count == 0:
                app.Quit()
        except Exception:
            pass
        shutil.rmtree(tmp.parent, ignore_errors=True)


def render_with_powerpoint(deck, out, width):
    def work(pres):
        ratio = pres.PageSetup.SlideHeight / pres.PageSetup.SlideWidth
        paths = []
        for i in range(1, pres.Slides.Count + 1):
            p = out / ("slide-%02d.png" % i)
            pres.Slides(i).Export(str(p), "PNG", width, round(width * ratio))
            paths.append(p)
        return paths
    return _with_powerpoint(deck, work)


def render_with_libreoffice(deck, out, width):
    import fitz                                                  # PyMuPDF
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RuntimeError("Neither PowerPoint (Windows) nor LibreOffice was found: export the slides from "
                           "PowerPoint (File > Export > PNG) and look at those.")
    tmpdir = pathlib.Path(tempfile.mkdtemp())
    profile = (tmpdir / "profile").as_uri()                      # its own profile: works even if LibreOffice is open
    subprocess.run([soffice, f"-env:UserInstallation={profile}", "--headless", "--convert-to", "pdf",
                    "--outdir", str(tmpdir), str(deck)], check=True)
    pdf = fitz.open(tmpdir / (deck.stem + ".pdf"))
    paths = []
    for i, page in enumerate(pdf, start=1):
        zoom = width / page.rect.width
        p = out / ("slide-%02d.png" % i)
        page.get_pixmap(matrix=fitz.Matrix(zoom, zoom)).save(str(p))
        paths.append(p)
    pdf.close()
    shutil.rmtree(tmpdir, ignore_errors=True)
    return paths


def render(deck, out, width=1600):
    deck, out = pathlib.Path(deck).resolve(), pathlib.Path(out).resolve()
    out.mkdir(parents=True, exist_ok=True)
    if os.name == "nt":
        try:
            return render_with_powerpoint(deck, out, width)
        except Exception as e:                                   # no pywin32 or no PowerPoint: try LibreOffice
            print(f"PowerPoint not available ({type(e).__name__}); trying LibreOffice.")
    return render_with_libreoffice(deck, out, width)


def check_fit(deck):
    """Windows with PowerPoint only. For every text box, compare the lines PowerPoint actually lays out with the lines
    the build intended (paragraphs plus line breaks), and flag text taller than its box. Returns a list of problems."""
    import pythoncom
    import win32com.client as win32

    def part(obj, name, *args):                                  # properties that take arguments need a direct call
        disp = obj._oleobj_
        r = disp.Invoke(disp.GetIDsOfNames(name), 0, pythoncom.DISPATCH_PROPERTYGET | pythoncom.DISPATCH_METHOD, True,
                        *args)
        return win32.Dispatch(r)

    def work(pres):
        problems = []
        for si in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(si)
            for shi in range(1, slide.Shapes.Count + 1):
                shp = slide.Shapes(shi)
                try:
                    if not shp.HasTextFrame:
                        continue
                    tr = shp.TextFrame2.TextRange
                    text = tr.Text
                except Exception:                                # locked shapes cannot be read
                    continue
                if not text.strip():
                    continue
                intended = part(tr, "Paragraphs").Count + text.count(chr(11))
                laid = 0
                while True:                                      # past the last line PowerPoint raises an error
                    try:
                        if part(tr, "Lines", laid + 1, 1).Length == 0:
                            break
                    except pythoncom.com_error:
                        break
                    laid += 1
                if laid != intended:
                    problems.append(f"slide {si} '{shp.Name}': PowerPoint lays out {laid} lines, the build set {intended}")
                if tr.BoundHeight > shp.Height + 1.5:
                    problems.append(f"slide {si} '{shp.Name}': text is {tr.BoundHeight - shp.Height:.1f} pt taller than its box")
        return problems
    return _with_powerpoint(deck, work)


if __name__ == "__main__":
    paths = render(sys.argv[1], sys.argv[2], int(sys.argv[3]) if len(sys.argv) > 3 else 1600)
    print("\n".join(str(p) for p in paths))
    if os.name == "nt":
        try:
            found = check_fit(sys.argv[1])
            print("\n".join(found) if found else "check_fit: every text box lays out as built")
        except Exception as e:
            print(f"check_fit skipped (needs PowerPoint and pywin32): {type(e).__name__}")
