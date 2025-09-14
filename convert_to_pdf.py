import os
import sys
import time
from pathlib import Path


SUPPORTED_EXTS = {".doc", ".docx", ".ppt", ".pptx"}


def find_office_files(src_root: Path):
    for path in src_root.rglob("*"):
        if path.is_file() and path.suffix.lower() in SUPPORTED_EXTS:
            yield path


def ensure_parent_dir(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)


def should_skip(src: Path, dst: Path) -> bool:
    if not dst.exists():
        return False
    try:
        return src.stat().st_mtime <= dst.stat().st_mtime
    except Exception:
        return False


def convert_with_win32com(files_by_app, dest_root: Path):
    """Convert using pywin32 COM automation if available."""
    try:
        import win32com.client  # type: ignore
        import pythoncom  # type: ignore
    except Exception as e:
        return False, f"win32com not available: {e}"

    # Initialize COM in STA
    pythoncom.CoInitialize()

    errors = []
    converted = 0

    try:
        # Word conversion
        word_files = files_by_app.get("word", [])
        word_app = None
        if word_files:
            try:
                word_app = win32com.client.DispatchEx("Word.Application")
                word_app.Visible = False
                word_app.DisplayAlerts = 0  # wdAlertsNone
            except Exception as e:
                errors.append(f"Failed to start Word: {e}")
                word_app = None

        if word_app is not None:
            wdFormatPDF = 17
            for src, dst in word_files:
                try:
                    ensure_parent_dir(dst)
                    if should_skip(src, dst):
                        continue
                    doc = word_app.Documents.Open(str(src), ReadOnly=True)
                    try:
                        # SaveAs2 works across versions, fall back to SaveAs when unavailable
                        try:
                            doc.SaveAs2(str(dst), FileFormat=wdFormatPDF)
                        except Exception:
                            doc.SaveAs(str(dst), FileFormat=wdFormatPDF)
                    finally:
                        doc.Close(False)
                    converted += 1
                except Exception as e:
                    errors.append(f"Word failed for {src}: {e}")

        # PowerPoint conversion
        ppt_files = files_by_app.get("ppt", [])
        ppt_app = None
        if ppt_files:
            try:
                ppt_app = win32com.client.DispatchEx("PowerPoint.Application")
                ppt_app.Visible = False
            except Exception as e:
                errors.append(f"Failed to start PowerPoint: {e}")
                ppt_app = None

        if ppt_app is not None:
            ppFixedFormatTypePDF = 2
            for src, dst in ppt_files:
                try:
                    ensure_parent_dir(dst)
                    if should_skip(src, dst):
                        continue
                    # WithWindow=False prevents opening a window
                    pres = ppt_app.Presentations.Open(str(src), WithWindow=False)
                    try:
                        pres.ExportAsFixedFormat(str(dst), ppFixedFormatTypePDF)
                    finally:
                        pres.Close()
                    converted += 1
                except Exception as e:
                    errors.append(f"PowerPoint failed for {src}: {e}")

    finally:
        try:
            # Quit apps if they were started
            try:
                word_app and word_app.Quit()
            except Exception:
                pass
            try:
                ppt_app and ppt_app.Quit()
            except Exception:
                pass
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    return True, (converted, errors)


def convert_with_powershell(files_by_app, dest_root: Path):
    """Fallback: invoke PowerShell COM automation per file."""
    import subprocess
    errors = []
    converted = 0

    # Build small PS functions for Word and PowerPoint
    ps_word = r'''
param($src,$dst)
$word = $null
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $doc = $word.Documents.Open($src, $true)
  try {
    $doc.SaveAs([ref]$dst, [ref]17)
  } finally {
    $doc.Close($false)
  }
} finally {
  if ($word -ne $null) { $word.Quit() | Out-Null }
}
'''

    ps_ppt = r'''
param($src,$dst)
$ppt = $null
try {
  $ppt = New-Object -ComObject PowerPoint.Application
  # Open as ReadOnly, no window
  $pres = $ppt.Presentations.Open($src, $true, $false, $false)
  try {
    # Use SaveAs with ppSaveAsPDF (32) for broader compatibility
    $pres.SaveAs($dst, 32)
  } finally {
    $pres.Close()
  }
} finally {
  if ($ppt -ne $null) { $ppt.Quit() | Out-Null }
}
'''

    for src, dst in files_by_app.get("word", []):
        try:
            ensure_parent_dir(dst)
            if should_skip(src, dst):
                continue
            cmd = [
                "powershell",
                "-NoProfile",
                "-Command",
                "function Convert-Word {" + ps_word + "}; Convert-Word -src '" + str(src).replace("'", "''") + "' -dst '" + str(dst).replace("'", "''") + "'",
            ]
            subprocess.run(cmd, check=True)
            converted += 1
        except Exception as e:
            errors.append(f"PowerShell Word failed for {src}: {e}")

    for src, dst in files_by_app.get("ppt", []):
        try:
            ensure_parent_dir(dst)
            if should_skip(src, dst):
                continue
            cmd = [
                "powershell",
                "-NoProfile",
                "-Command",
                "function Convert-Ppt {" + ps_ppt + "}; Convert-Ppt -src '" + str(src).replace("'", "''") + "' -dst '" + str(dst).replace("'", "''") + "'",
            ]
            subprocess.run(cmd, check=True)
            if dst.exists():
                converted += 1
            else:
                errors.append(f"PowerShell PowerPoint produced no file for {src}")
        except Exception as e:
            errors.append(f"PowerShell PowerPoint failed for {src}: {e}")

    return True, (converted, errors)


def main():
    cwd = Path(os.getcwd())
    src_root = cwd / "To Change"
    dest_root = cwd / "PDFs"

    if not src_root.exists():
        print(f"Source folder not found: {src_root}")
        return 2
    dest_root.mkdir(parents=True, exist_ok=True)

    # Prepare mapping for apps
    files_by_app = {"word": [], "ppt": []}

    for src in find_office_files(src_root):
        rel = src.relative_to(src_root)
        dst = dest_root / rel
        dst = dst.with_suffix(".pdf")
        if src.suffix.lower() in {".doc", ".docx"}:
            files_by_app["word"].append((src, dst))
        elif src.suffix.lower() in {".ppt", ".pptx"}:
            files_by_app["ppt"].append((src, dst))

    total_files = len(files_by_app["word"]) + len(files_by_app["ppt"])
    if total_files == 0:
        print("No Office files found to convert.")
        return 0

    print(f"Found {total_files} files. Starting conversion...")

    # Try win32com first
    ok, result = convert_with_win32com(files_by_app, dest_root)
    if ok:
        converted, errors = result
    else:
        print(str(result))
        print("Falling back to PowerShell-based conversion...")
        ok, result = convert_with_powershell(files_by_app, dest_root)
        converted, errors = result if ok else (0, [str(result)])

    print(f"Converted: {converted} / {total_files}")
    if errors:
        print("Errors:")
        for e in errors:
            print(" - ", e)
        # Return non-zero if any errors occurred
        return 1 if converted > 0 else 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
