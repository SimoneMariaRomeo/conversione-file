sapijdpais jpo sda'0 is'da0i'0dsi'0aisd'0 ia 

import os
from pathlib import Path
import subprocess


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


def _ppt_collect_text_from_shapes(shapes) -> list[str]:
    """Recursively collect visible text from PowerPoint shapes via COM."""
    lines: list[str] = []
    try:
        count = shapes.Count
    except Exception:
        return lines

    # COM collections are 1-indexed
    for i in range(1, int(count) + 1):
        try:
            shape = shapes.Item(i)
        except Exception:
            continue

        # Grouped shapes (msoGroup == 6)
        try:
            if int(getattr(shape, "Type", 0)) == 6:
                try:
                    group_items = shape.GroupItems
                    lines.extend(_ppt_collect_text_from_shapes(group_items))
                except Exception:
                    pass
        except Exception:
            pass

        # Tables
        try:
            if int(getattr(shape, "HasTable", 0)) == -1:
                try:
                    rows = shape.Table.Rows.Count
                    cols = shape.Table.Columns.Count
                    for r in range(1, int(rows) + 1):
                        for c in range(1, int(cols) + 1):
                            try:
                                cell = shape.Table.Cell(r, c)
                                text = cell.Shape.TextFrame.TextRange.Text
                                if text and str(text).strip():
                                    lines.append(str(text))
                            except Exception:
                                pass
                except Exception:
                    pass
        except Exception:
            pass

        # Text frames
        try:
            if int(getattr(shape, "HasTextFrame", 0)) == -1:
                try:
                    if int(shape.TextFrame.HasText) == -1:
                        text = shape.TextFrame.TextRange.Text
                        if text and str(text).strip():
                            lines.append(str(text))
                except Exception:
                    pass
        except Exception:
            pass

    return lines


def convert_with_win32com(files_by_app, dest_root: Path):
    """Convert using pywin32 COM automation if available."""
    try:
        import win32com.client  # type: ignore
        import pythoncom  # type: ignore
    except Exception as e:
        return False, f"win32com not available: {e}"

    # Initialize COM in STA
    pythoncom.CoInitialize()

    errors: list[str] = []
    converted = 0

    try:
        # Word conversion to Unicode text
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
            wdFormatUnicodeText = 7
            for src, dst in word_files:
                try:
                    ensure_parent_dir(dst)
                    if should_skip(src, dst):
                        continue
                    doc = word_app.Documents.Open(str(src), ReadOnly=True)
                    try:
                        try:
                            doc.SaveAs2(str(dst), FileFormat=wdFormatUnicodeText)
                        except Exception:
                            doc.SaveAs(str(dst), FileFormat=wdFormatUnicodeText)
                    finally:
                        doc.Close(False)
                    converted += 1
                except Exception as e:
                    errors.append(f"Word failed for {src}: {e}")

        # PowerPoint conversion: extract text content
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
            for src, dst in ppt_files:
                try:
                    ensure_parent_dir(dst)
                    if should_skip(src, dst):
                        continue
                    pres = ppt_app.Presentations.Open(str(src), WithWindow=False)
                    try:
                        lines: list[str] = []
                        try:
                            slide_count = int(pres.Slides.Count)
                        except Exception:
                            slide_count = 0
                        for idx in range(1, slide_count + 1):
                            try:
                                slide = pres.Slides.Item(idx)
                                lines.append(f"Slide {idx}")
                                lines.extend(_ppt_collect_text_from_shapes(slide.Shapes))
                                lines.append("")
                            except Exception:
                                pass
                        dst.write_text("\n".join(lines), encoding="utf-8")
                        converted += 1
                    finally:
                        pres.Close()
                except Exception as e:
                    errors.append(f"PowerPoint failed for {src}: {e}")

    finally:
        try:
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
    errors: list[str] = []
    converted = 0

    # Word to Unicode text
    ps_word = r'''
param($src,$dst)
$word = $null
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $doc = $word.Documents.Open($src, $true)
  try {
    # 7 = wdFormatUnicodeText
    $doc.SaveAs([ref]$dst, [ref]7)
  } finally {
    $doc.Close($false)
  }
} finally {
  if ($word -ne $null) { $word.Quit() | Out-Null }
}
'''

    # PowerPoint text extraction
    ps_ppt = r'''
param($src,$dst)
$ppt = $null
try {
  $ppt = New-Object -ComObject PowerPoint.Application
  $pres = $ppt.Presentations.Open($src, $true, $false, $false)
  try {
    $sb = New-Object System.Text.StringBuilder
    $slideCount = [int]$pres.Slides.Count
    for ($i = 1; $i -le $slideCount; $i++) {
      $slide = $pres.Slides.Item($i)
      [void]$sb.AppendLine("Slide $i")
      $shapes = $slide.Shapes
      $shapeCount = [int]$shapes.Count
      for ($s = 1; $s -le $shapeCount; $s++) {
        $shape = $shapes.Item($s)

        # Grouped shapes (msoGroup == 6)
        if ($shape.Type -eq 6) {
          $groupItems = $shape.GroupItems
          $gCount = [int]$groupItems.Count
          for ($g = 1; $g -le $gCount; $g++) {
            $gshape = $groupItems.Item($g)
            if ($gshape.HasTextFrame -and $gshape.TextFrame.HasText) {
              $t = $gshape.TextFrame.TextRange.Text
              if ($t -and $t.Trim().Length -gt 0) { [void]$sb.AppendLine($t) }
            }
          }
        }

        if ($shape.HasTable) {
          $rows = [int]$shape.Table.Rows.Count
          $cols = [int]$shape.Table.Columns.Count
          for ($r = 1; $r -le $rows; $r++) {
            for ($c = 1; $c -le $cols; $c++) {
              $cell = $shape.Table.Cell($r,$c)
              $ct = $cell.Shape.TextFrame.TextRange.Text
              if ($ct -and $ct.Trim().Length -gt 0) { [void]$sb.AppendLine($ct) }
            }
          }
        }

        if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
          $t = $shape.TextFrame.TextRange.Text
          if ($t -and $t.Trim().Length -gt 0) { [void]$sb.AppendLine($t) }
        }
      }
      [void]$sb.AppendLine("")
    }
    [System.IO.File]::WriteAllText($dst, $sb.ToString(), [System.Text.Encoding]::UTF8)
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
            if dst.exists():
                converted += 1
            else:
                errors.append(f"PowerShell Word produced no file for {src}")
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
    dest_root = cwd / "TXT"

    if not src_root.exists():
        print(f"Source folder not found: {src_root}")
        return 2
    dest_root.mkdir(parents=True, exist_ok=True)

    files_by_app = {"word": [], "ppt": []}

    for src in find_office_files(src_root):
        rel = src.relative_to(src_root)
        dst = dest_root / rel
        dst = dst.with_suffix(".txt")
        if src.suffix.lower() in {".doc", ".docx"}:
            files_by_app["word"].append((src, dst))
        elif src.suffix.lower() in {".ppt", ".pptx"}:
            files_by_app["ppt"].append((src, dst))

    total_files = len(files_by_app["word"]) + len(files_by_app["ppt"])
    if total_files == 0:
        print("No Office files found to convert.")
        return 0

    print(f"Found {total_files} files. Starting conversion to TXT...")

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
        return 1 if converted > 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
