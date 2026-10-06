"""見積ファイル（Excel）の指定シートを、LibreOffice で1つのPDFに変換する。

smm-review（画面側リポジトリ）の server/convert.py（Windows + Excel COM 版）を、
Cloud Run（Linux）で動くように LibreOffice（UNO）で作り直したもの。
処理の流れ（パターンA/Bの判定、対象シートだけを出力、パターンBの印刷設定の補正）は同じ。

1回の変換ごとに LibreOffice を専用のユーザープロファイルで起動し、終わったら終了する
（変換同士が干渉しないようにするため）。main.py からは別プロセスとして呼び出す：

    python3 convert.py <入力.xlsx> <出力.pdf>
    → 標準出力に {"pattern": "A", "warnings": [...]} を JSON で出す
"""

from __future__ import annotations

import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import time
import uuid

PATTERN_A = ("御見積書", "割付図・系統図", "架台図", "発電シミュレーション")
PATTERN_B = ("見積書", "配置図", "発電シミュレーション")
EXCEL_EXTENSIONS = {".xls", ".xlsx", ".xlsm"}
FALLBACK_PRINT_AREAS = {
    "見積書": "$A$1:$M$103",
    "配置図": "$A$1:$BP$44,$A$46:$BP$85",
    "発電シミュレーション": "$A$1:$BS$62",
}
# パターンBの用紙の向き（Excel版と同じ：見積書は縦、それ以外は横）
PORTRAIT_SHEETS = {"見積書"}
A4 = (21000, 29700)  # 1/100 mm

SOFFICE = os.environ.get("SOFFICE", "soffice")
# Docker イメージのビルド時に作っておく LibreOffice のプロファイル（毎回の初期化を省いて起動を速くする）
PROFILE_TEMPLATE = Path(os.environ.get("LO_PROFILE_TEMPLATE", "/opt/lo-profile"))
CONNECT_TIMEOUT = 45


def detect_pattern(visible: list[str]) -> tuple[str, list[str]]:
    if all(name in visible for name in PATTERN_A):
        return "A", list(PATTERN_A)
    if all(name in visible for name in PATTERN_B):
        return "B", list(PATTERN_B)
    raise ValueError("パターンAまたはBに必要なシートが揃っていません。")


def _props(**kwargs):
    from com.sun.star.beans import PropertyValue  # type: ignore  # LibreOffice 付属（Cloud Run のコンテナ内にだけある）

    return tuple(PropertyValue(Name=k, Value=v) for k, v in kwargs.items())


class Office:
    """変換1回分の LibreOffice（headless）を起動し、UNO で接続する。"""

    def __enter__(self):
        import uno  # type: ignore  # LibreOffice 付属（Cloud Run のコンテナ内にだけある）

        self.tmp = Path(tempfile.mkdtemp(prefix="lo_"))
        profile = self.tmp / "profile"
        if PROFILE_TEMPLATE.is_dir():
            shutil.copytree(PROFILE_TEMPLATE, profile)
        pipe = "smm" + uuid.uuid4().hex
        self.proc = subprocess.Popen(
            [
                SOFFICE, "--headless", "--invisible", "--nologo", "--norestore",
                "--nodefault", "--nolockcheck", "--nofirststartwizard",
                f"-env:UserInstallation={uno.systemPathToFileUrl(str(profile))}",
                f"--accept=pipe,name={pipe};urp;StarOffice.ComponentContext",
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
        deadline = time.monotonic() + CONNECT_TIMEOUT
        while True:
            try:
                ctx = resolver.resolve(f"uno:pipe,name={pipe};urp;StarOffice.ComponentContext")
                break
            except Exception:
                if self.proc.poll() is not None or time.monotonic() > deadline:
                    self.__exit__(None, None, None)
                    raise RuntimeError("LibreOfficeを起動できませんでした。")
                time.sleep(0.2)
        self.desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        return self

    def __exit__(self, *exc):
        try:
            if getattr(self, "desktop", None) is not None:
                self.desktop.terminate()
        except Exception:
            pass  # terminate 中に接続が切れるのは正常
        try:
            self.proc.wait(timeout=10)
        except Exception:
            self.proc.kill()
        shutil.rmtree(self.tmp, ignore_errors=True)


def _parse_ranges(sheet, spec: str):
    return tuple(
        sheet.getCellRangeByName(part.replace("$", "")).getRangeAddress()
        for part in spec.split(",")
    )


def _own_page_style(doc, sheet, used_count: dict[str, int]):
    """シート専用のページスタイルを返す（他のシートと共有していれば複製してから割り当てる）。"""
    styles = doc.StyleFamilies.getByName("PageStyles")
    name = sheet.PageStyle
    if used_count.get(name, 0) <= 1:
        return styles.getByName(name)
    original = styles.getByName(name)
    new_name = f"smm_{sheet.Name}"
    style = doc.createInstance("com.sun.star.style.PageStyle")
    styles.insertByName(new_name, style)
    for prop in original.PropertySetInfo.Properties:
        try:
            style.setPropertyValue(prop.Name, original.getPropertyValue(prop.Name))
        except Exception:
            pass  # 読み取り専用などコピーできない項目は既定値のまま
    sheet.PageStyle = new_name
    return style


def _auto_layout_charts(sheet) -> None:
    """グラフの描画領域（プロットエリア）の位置・大きさを自動に戻す。

    パターンBの古い .xls のグラフは、Excel 側で手動指定されたプロットエリアの配置を LibreOffice が
    正しく読めず、幅がほぼ0に潰れて空のグラフに見えるため。
    """
    charts = sheet.getCharts()
    for name in charts.getElementNames():
        try:
            model = charts.getByName(name).getEmbeddedObject()
            diagram = model.getDiagram()
        except Exception:
            continue
        try:
            diagram.setAutomaticDiagramPositioning()
        except Exception:
            pass
        try:  # chart2 側の手動配置（相対位置・相対サイズ）も既定に戻す
            first = model.getFirstDiagram()
            for prop in ("RelativePosition", "RelativeSize"):
                try:
                    first.setPropertyToDefault(prop)
                except Exception:
                    pass
        except Exception:
            pass
        try:
            model.setModified(True)
        except Exception:
            pass


def _strip_trailing_newlines(sheet) -> None:
    """文字列セルの末尾の改行を取り除く。

    Excel は「本文＋末尾の改行」のセルを1行として表示するが、LibreOffice は空行を含む2行として
    扱うため、文字が上にずれて低い行からはみ出し、半分隠れる（パターンBの発電シミュレーション B61 で確認）。
    数式のセルは対象外（定数の文字列セルだけ）。
    """
    string_cells = 4  # com.sun.star.sheet.CellFlags.STRING
    try:
        cells = sheet.queryContentCells(string_cells).getCells().createEnumeration()
    except Exception:
        return
    while cells.hasMoreElements():
        cell = cells.nextElement()
        text = cell.getString()
        if text.endswith(("\n", "\r")):
            cell.setString(text.rstrip("\r\n"))


def convert_auto(source: Path, destination: Path) -> tuple[str, list[str]]:
    import uno  # type: ignore  # LibreOffice 付属（Cloud Run のコンテナ内にだけある）

    if source.suffix.lower() not in EXCEL_EXTENSIONS:
        raise ValueError(".xls / .xlsx / .xlsm ファイルを指定してください。")
    warnings: list[str] = []
    with Office() as office:
        doc = office.desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(str(source)), "_blank", 0,
            _props(Hidden=True, UpdateDocMode=0, MacroExecutionMode=0),
        )
        if doc is None or not hasattr(doc, "Sheets"):
            raise ValueError("Excelファイルとして開けませんでした。")
        try:
            sheets = [doc.Sheets.getByIndex(i) for i in range(doc.Sheets.Count)]
            visible = [s.Name for s in sheets if s.IsVisible]
            pattern, names = detect_pattern(visible)
            used_count: dict[str, int] = {}
            for s in sheets:
                used_count[s.PageStyle] = used_count.get(s.PageStyle, 0) + 1

            for s in sheets:
                if s.Name not in names:
                    continue
                if not s.getPrintAreas() and s.Name in FALLBACK_PRINT_AREAS:
                    s.setPrintAreas(_parse_ranges(s, FALLBACK_PRINT_AREAS[s.Name]))
                    warnings.append(f"{s.Name}: 印刷範囲がなかったため、標準範囲を設定しました。")
                if not s.getPrintAreas():
                    warnings.append(f"{s.Name}: 印刷範囲が未設定です。使用範囲を出力します。")
                _strip_trailing_newlines(s)
                if pattern == "B":
                    # パターンBの古いテンプレートは印刷設定が不安定なため、A4・1ページ（印刷範囲ごと）に揃える
                    style = _own_page_style(doc, s, used_count)
                    portrait = s.Name in PORTRAIT_SHEETS
                    style.IsLandscape = not portrait
                    style.Width, style.Height = A4 if portrait else (A4[1], A4[0])
                    style.ScaleToPagesX = 1
                    style.ScaleToPagesY = max(1, len(s.getPrintAreas()))
                    # パターンBの発電シミュレーションのグラフは、そのままだとプロットエリアが潰れて空に見える
                    _auto_layout_charts(s)

            # 対象外のシートは出力から外す（変更はメモリ上だけで、元ファイルは保存しない）。
            # LibreOffice は非表示のシートでも印刷範囲があると PDF に出力するため、非表示にしたうえで印刷範囲も外す。
            try:  # 表示中のシートは非表示にできないため、先に対象シートを表示中にしておく
                doc.CurrentController.setActiveSheet(next(s for s in sheets if s.Name == names[0]))
            except Exception:
                pass
            for s in sheets:
                if s.Name in names:
                    continue
                if s.IsVisible:
                    s.IsVisible = False
                if s.getPrintAreas():
                    s.setPrintAreas(())

            doc.storeToURL(
                uno.systemPathToFileUrl(str(destination)),
                _props(FilterName="calc_pdf_Export"),
            )
        finally:
            try:
                doc.close(True)
            except Exception:
                pass
    data = destination.read_bytes() if destination.is_file() else b""
    if len(data) < 100 or not data.startswith(b"%PDF-"):
        raise RuntimeError("PDFを作成できませんでした。")
    return pattern, warnings


def main() -> int:
    if len(sys.argv) != 3:
        print("使い方: python3 convert.py <入力.xlsx> <出力.pdf>", file=sys.stderr)
        return 2
    try:
        pattern, warnings = convert_auto(Path(sys.argv[1]).resolve(), Path(sys.argv[2]).resolve())
    except ValueError as exc:  # 入力ファイルの問題（利用者に見せてよいメッセージ）
        print(json.dumps({"error": str(exc), "kind": "input"}, ensure_ascii=False))
        return 1
    except Exception as exc:
        print(json.dumps({"error": str(exc), "kind": "server"}, ensure_ascii=False))
        return 1
    print(json.dumps({"pattern": pattern, "warnings": warnings}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
