from pathlib import Path

path = Path('main.py')
text = path.read_text(encoding='utf-8')
start = text.index("        fallback_paths = [")
end = text.index("        for path in fallback_paths:", start)
new_block = """        fallback_paths = [\n            Path(\"C:/Windows/Fonts/yu-gothic.ttf\"),\n            Path(\"C:/Windows/Fonts/yu-gothic.ttc\"),\n            Path(\"C:/Windows/Fonts/YuGothR.ttc\"),\n            Path(\"C:/Windows/Fonts/meiryo.ttc\"),\n            Path(\"C:/Windows/Fonts/meiryob.ttc\"),\n            Path(\"C:/Windows/Fonts/msgothic.ttc\"),\n            Path(\"/System/Library/Fonts/ヒラギノ角ゴシック W4.ttc\"),\n            Path(\"/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc\"),\n        ]\n"""
text = text[:start] + new_block + text[end:]
path.write_text(text, encoding='utf-8')
