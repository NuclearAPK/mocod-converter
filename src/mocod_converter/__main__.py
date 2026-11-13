import argparse
from pathlib import Path
from zipfile import BadZipFile

from openpyxl.utils.exceptions import InvalidFileException

from .converter import ConversionError, convert


def main():
    parser = argparse.ArgumentParser(description="Конвертер реестра займов МФО для МоКоД")
    parser.add_argument("input", nargs="?", type=Path, help="исходный реестр XLSX")
    parser.add_argument("output", nargs="?", type=Path, help="путь к выходному CSV")
    parser.add_argument("-f", "--force", action="store_true", help="заменить существующий CSV")
    args = parser.parse_args()
    if args.input is None and args.output is None and not args.force:
        try:
            import tkinter

            from .gui import main as run_gui
        except ImportError:
            parser.exit(1, "Для графического режима установите Tk. CLI работает без Tk.\n")
        try:
            run_gui()
        except tkinter.TclError:
            parser.exit(1, "Не удалось открыть окно. Проверьте Tk и доступность рабочего стола.\n")
        return
    if args.input is None or args.output is None:
        parser.error("Укажите исходный XLSX и путь к выходному CSV.")
    if args.output.exists() and not args.force:
        parser.exit(1, "Выходной файл уже существует. Для замены укажите --force.\n")
    try:
        count = convert(args.input, args.output)
    except ConversionError as error:
        parser.exit(1, f"{error}\n")
    except (BadZipFile, InvalidFileException, OSError):
        parser.exit(1, "Не удалось прочитать XLSX или записать CSV. Проверьте файл и папку назначения.\n")
    print(f"CSV сохранён. Строк в выгрузке — {count}.")


if __name__ == "__main__":
    main()
