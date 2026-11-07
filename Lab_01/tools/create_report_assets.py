"""Capture real command output as text and legible console-style images."""
from pathlib import Path
import os
import subprocess
import sys
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parents[1]


def capture(args: list[str], stdin: str | None = None) -> str:
    env = dict(os.environ, PYTHONIOENCODING='utf-8')
    result = subprocess.run(args, cwd=ROOT, input=stdin, capture_output=True,
                            text=True, encoding='utf-8', env=env)
    if result.returncode:
        raise RuntimeError(result.stdout + result.stderr)
    return result.stdout + result.stderr


def image(name: str, command: str, output: str) -> None:
    folder = ROOT / 'docs/screenshots'
    folder.mkdir(parents=True, exist_ok=True)
    (folder / (name + '.txt')).write_text(command + '\n' + output, encoding='utf-8')
    font = ImageFont.truetype('C:/Windows/Fonts/consola.ttf', 20)
    lines = (command + '\n\n' + output).expandtabs(4).splitlines()
    width = max(1000, int(max(font.getlength(line) for line in lines)) + 48)
    canvas = Image.new('RGB', (width, len(lines) * 27 + 48), '#16202b')
    draw = ImageDraw.Draw(canvas)
    for index, line in enumerate(lines):
        draw.text((24, 24 + index * 27), line, font=font, fill='#edf2f7')
    canvas.save(folder / (name + '.png'))


def main() -> None:
    python = str(ROOT / '.venv/Scripts/python.exe')
    if not Path(python).exists():
        python = sys.executable
    demo = capture([python, '-m', 'book_manager'])
    lines = demo.splitlines()
    image('application_output_1', 'python -m book_manager', '\n'.join(lines[:12]))
    image('application_output_2', 'python -m book_manager (продовження)',
          '\n'.join(lines[12:]))
    image('test_results', 'python -m unittest discover -s tests -v',
          capture([python, '-m', 'unittest', 'discover', '-s', 'tests', '-v']))
    image('git_status', 'git status', capture(['git', 'status']))
    image('git_log', 'git log --oneline', capture(['git', 'log', '--oneline']))


if __name__ == '__main__':
    main()
