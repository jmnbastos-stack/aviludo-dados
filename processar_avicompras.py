import xlrd, json, os, sys, requests, base64, glob

GITHUB_TOKEN = os.environ["GH_TOKEN"]
GITHUB_REPO = "jmnbastos-stack/aviludo-dados"
headers = {"Authorization": f"token {GITHUB_TOKEN}", "Accept": "application/vnd.github.v3+json"}

# Encontrar ficheiro AVI-COMPRAS
ficheiros = glob.glob("AVI-COMPRAS*.xls") + glob.glob("AVI-COMPRAS*.xlsx")
if not ficheiros:
    print("Nenhum ficheiro AVI-COMPRAS encontrado!")
    sys.exit(1)

ficheiro = sorted(ficheiros)[-1]
print(f"A processar: {ficheiro}")

# Mapeamento armazéns (ignorar SA, SP, BJ)
ARM_MAP = {
    "LX": {"cod": 10, "qt": 12, "pmp": 23},
    "LS": {"cod": 25, "qt": 27, "pmp": 38},
    "PT": {"cod": 85, "qt": 87, "pmp": 98},
    "SD": {"cod": 100, "qt": 102, "pmp": 113},
}

def safe_float(v):
    try:
        s = str(v).replace(",", ".").replace("\xa0", "").replace(" ", "").replace("\u202f", "")
        f = float(s)
        import math
        return round(f, 4) if not math.isnan(f) else 0.0
    except:
        return 0.0

wb = xlrd.open_workbook(ficheiro)
sh = wb.sheet_by_index(0)
print(f"Linhas: {sh.nrows}, Colunas: {sh.ncols}")

artigos = []
for i in range(2, sh.nrows):
    cod_raw = sh.cell_value(i, 0)
    if not cod_raw: continue
    try:
        cod = str(int(float(str(cod_raw)))).strip()
    except:
        cod = str(cod_raw).strip().split(".")[0]
    if not cod or not cod.isdigit(): continue

    desc = str(sh.cell_value(i, 1)).strip()
    status = str(sh.cell_value(i, 5)).strip()
    un = str(sh.cell_value(i, 9)).strip()

    stock = {}
    for arm, cols in ARM_MAP.items():
        qt = safe_float(sh.cell_value(i, cols["qt"]))
        pmp = safe_float(sh.cell_value(i, cols["pmp"]))
        stock[arm] = {"qt_disp": qt, "pmp": pmp}

    artigos.append({"cod": cod, "desc": desc, "status": status, "un": un, "stock": stock})

print(f"Artigos processados: {len(artigos)}")

# Upload para GitHub
content = json.dumps(artigos, ensure_ascii=False)
b64 = base64.b64encode(content.encode()).decode()

url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/artigos.json"
r = requests.get(url, headers=headers)
sha = r.json().get("sha", "") if r.status_code == 200 else ""

payload = {"message": f"artigos.json actualizado automaticamente - {ficheiro}", "content": b64}
if sha: payload["sha"] = sha

r2 = requests.put(url, headers=headers, json=payload)
if r2.status_code in (200, 201):
    print("✅ artigos.json actualizado no GitHub!")
else:
    print(f"❌ Erro: {r2.json().get('message', '')}")
    sys.exit(1)
