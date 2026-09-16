from html import escape


DEFAULT_STYLE = """
body { font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #222; }
h1 { font-size: 18pt; color: #1f4e79; }
h2 { font-size: 14pt; color: #2f5597; margin-top: 18px; }
p { margin: 8px 0; line-height: 1.4; }
table { border-collapse: collapse; width: 100%; margin: 12px 0; }
th { background: #1f4e79; color: white; padding: 8px; text-align: left; }
td { border: 1px solid #d9d9d9; padding: 8px; }
li { margin: 4px 0; }
"""


def markdown_to_html(markdown: str) -> str:
    lines = markdown.replace("\r\n", "\n").split("\n")
    html = ["<html><head><style>", DEFAULT_STYLE, "</style></head><body>"]
    in_list = False

    def close_list():
        nonlocal in_list
        if in_list:
            html.append("</ul>")
            in_list = False

    i = 0
    while i < len(lines):
        line = lines[i]
        if line.startswith("# "):
            close_list()
            html.append(f"<h1>{escape(line[2:])}</h1>")
        elif line.startswith("## "):
            close_list()
            html.append(f"<h2>{escape(line[3:])}</h2>")
        elif line.startswith("- "):
            if not in_list:
                html.append("<ul>")
                in_list = True
            html.append(f"<li>{escape(line[2:])}</li>")
        elif line.strip() == "":
            close_list()
        else:
            close_list()
            html.append(f"<p>{escape(line)}</p>")
        i += 1

    close_list()
    html.append("</body></html>")
    return "".join(html)
