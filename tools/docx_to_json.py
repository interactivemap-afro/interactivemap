#!/usr/bin/env python3
"""
Converte fichas DOCX (formato "NOMBRE: ... / Área actuación: ... / Resumen: ...")
em registos compatíveis com o data.json do mapa CEsA.

Uso:
  python3 tools/docx_to_json.py --data data.json --meta tools/fichas_meta_es.json fichas_es/*.docx

- Os registos existentes NÃO são alterados. Os novos entram com "projeto": "es"
  e IDs a partir de 1001 (para não colidir com os IDs portugueses).
- Um registo já importado (mesmo nome) é atualizado em vez de duplicado (mantém o ID).
- tools/fichas_meta_es.json (chave = nome do ficheiro .docx, sem extensão):
    { "coords": [lat, lon], "logo": "LogoX.jpg", "evento": true,
      "nome": "Nome a mostrar", "categoria": "slug", "excluirLinks": ["url..."] }
  Tudo opcional. "logo" é procurado na pasta dos .docx; sem logo usa a imagem embutida.
- Registos sem morada passam a "digital" (aparecem na lista Mundo Digital).
"""
import argparse, json, re, shutil, sys, unicodedata, zipfile
from pathlib import Path
from docx import Document

ID_START = 1001

# Forma jurídica (ES) -> slug "categoria" (tem de existir icons/<slug>.svg e icons_filter/<slug>.svg)
# Usa-se a primeira parte (separada por "/") que tiver correspondência.
FORMA_MAP = {
    "asociacion": "associacao", "associacio": "associacao",
    "cooperativa": "cooperativa",
    "empresa": "empresa", "agencia": "empresa",
    "colectivo": "coletivo", "colectivo cultural": "coletivo_cultural",
    "grupo": "grupo", "grupo teatral": "grupo_teatral",
    "compania de teatro": "grupo_teatral", "compania": "companhia",
    "centro": "centro", "red": "rede", "fundacion": "instituto",
    "movimiento": "movimento", "federacion": "uniao", "union": "uniao",
    "frente": "frente", "liga": "liga", "instituto": "instituto",
    "espacio comunitario": "espaco_comunitario",
    "blog": "medio", "radio": "medio", "podcast": "medio",
}

# Área de actuación (ES) -> slug subCategoria (só as que têm ícone)
AREA_MAP = {
    "promocion cultural": "promocao_cultural", "educacion": "educacao",
    "musica": "musica", "audiovisual": "audiovisual", "danza": "danca",
    "teatro": "teatro", "grupo teatral": "teatro", "deporte": "desporto",
    "gastronomia": "gastronomia",
}

SOCIAL = [("instagram.com", "instagram"), ("facebook.com", "facebook"), ("fb.com", "facebook"),
          ("youtube.com", "youtube"), ("youtu.be", "youtube"), ("x.com", "twitter"),
          ("twitter.com", "twitter"), ("linkedin.com", "linkedin"), ("tiktok.com", "tiktok"),
          ("threads.net", "thread")]

LABELS = {
    "nombre": "nome", "area actuacion": "area", "area de actuacion": "area",
    "forma juridica": "forma", "resumen": "sobre", "resumen ingles": "sobre_en",
    "resumen portugues": "sobre_pt", "direccion": "direccion", "direccion postal": "direccion",
    "telefono": "telefone", "email": "email", "e-mail": "email", "correo": "email",
    "website": "website", "web": "website", "website ultimo festival": "website",
    "actividades": "atividades",
}
TEXT_KEYS = ("sobre", "sobre_en", "sobre_pt", "atividades")
EMPTY = {"", "-", "–", "—", "no tiene", "no tienen", "n/a"}


def norm(s):
    s = unicodedata.normalize("NFD", s or "").encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip().lower()


def slug(s):
    return re.sub(r"[^a-z0-9]+", "-", norm(s)).strip("-")


def clean_url(u):
    u = re.split(r"\)|%20|\s", u.strip())[0].rstrip(".,;\"'")
    u = u.replace("/watch&v=", "/watch?v=")
    u = re.sub(r"(twitter\.com|x\.com)/@", r"\1/", u)
    u = re.sub(r"[?&](utm_[^&]*|igsh=[^&]*|si=[^&]*|themeRefresh=[^&]*|cbrd=[^&]*)", "", u)
    u = re.sub(r"[?&]feature=[^&]*", "", u)
    u = re.sub(r"\?&", "?", u).rstrip("?&")
    if "&" in u and "?" not in u:          # "watch&v=..." -> "watch?v=..."
        u = u.replace("&", "?", 1)
    return u


def match_label(p):
    """Devolve (chave, valor) se o parágrafo começar por um rótulo conhecido (com ou sem ':')."""
    m = re.match(r"^\s*([^:\n]{2,40}?)\s*:\s*(.*)$", p, re.S)
    if m and LABELS.get(norm(m.group(1))):
        return LABELS[norm(m.group(1))], m.group(2)
    # rótulos sem dois pontos: "Teléfono 608748873", "Email  x@y", "Dirección -"
    n = norm(p)
    for lab in sorted(LABELS, key=len, reverse=True):
        if n == lab or n.startswith(lab + " "):
            words = len(lab.split())
            rest = re.split(r"\s+", p.strip(), maxsplit=words)
            return LABELS[lab], rest[words] if len(rest) > words else ""
    return None, None


def read_docx(path):
    doc = Document(path)
    raw = [p.text for p in doc.paragraphs]
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells:
                raw.append(c.text)
    # quebras de linha dentro do mesmo parágrafo contam como linhas separadas
    paras = [l.replace("​", "").replace("\xa0", " ").strip()
             for r in raw for l in re.split(r"[\n\v]", r)]
    urls = []
    for rel in doc.part.rels.values():
        if "hyperlink" in rel.reltype and rel.is_external:
            urls.append(rel.target_ref)

    fields, last, social_lines = {}, None, []
    for p in paras:
        if not p:
            continue
        key, val = match_label(p)
        if key and key not in fields:
            fields[key] = val.strip()
            last = key
        elif key:
            last = None
        elif last in TEXT_KEYS and not re.match(r"^\s*https?://\S+\s*$", p):
            fields[last] += "\n" + p
        else:
            if re.search(r"https?://|@\w", p):
                social_lines.append(p)
            last = None if last not in TEXT_KEYS else last
    text = "\n".join(social_lines)
    urls += re.findall(r"https?://[^\s)\"]+", text)
    # handles "in (@conta)" / "Youtube (@Canal)"
    for net, h in re.findall(r"\b(in|instagram|youtube|x|tiktok)\s*\(\s*@([\w.]+)\s*\)", text, re.I):
        net = net.lower()
        base = {"in": "https://www.instagram.com/", "instagram": "https://www.instagram.com/",
                "youtube": "https://www.youtube.com/@", "x": "https://x.com/",
                "tiktok": "https://www.tiktok.com/@"}[net]
        urls.append(base + h)
    for k in list(fields):
        v = fields[k].strip().rstrip(".").strip()
        fields[k] = "" if norm(v) in EMPTY else v
    return fields, [clean_url(u) for u in urls]


def detect_ext(b):
    return ("png" if b[:8] == b"\x89PNG\r\n\x1a\n" else
            "jpg" if b[:3] == b"\xff\xd8\xff" else
            "gif" if b[:4] == b"GIF8" else
            "webp" if b[8:12] == b"WEBP" else None)


def save_image(b, out_dir, name):
    ext = detect_ext(b)
    if not ext:
        return ""
    out_dir.mkdir(parents=True, exist_ok=True)
    for old in out_dir.glob(f"{slug(name)}.*"):
        old.unlink()
    dest = out_dir / f"{slug(name)}.{ext}"
    dest.write_bytes(b)
    return dest


def extract_image(path, out_dir, name, logo_file=None):
    if logo_file and Path(logo_file).exists():
        return save_image(Path(logo_file).read_bytes(), out_dir, name)
    with zipfile.ZipFile(path) as z:
        found = [(i.file_size, z.read(i)) for i in z.infolist() if i.filename.startswith("word/media/")]
        found = [(s, b) for s, b in found if detect_ext(b)]
        if not found:
            return ""
        return save_image(max(found)[1], out_dir, name)


def parse_address(raw):
    m = re.search(r"\b(\d{5})\b,?\s*(.*)$", raw or "")
    if m:
        street = raw[:m.start()].strip(" ,")
        city = m.group(2).strip(" ,.") or ""
        if city.isupper():
            city = city.title()
        return street, m.group(1), city
    parts = [x.strip() for x in (raw or "").split(",")]
    return (", ".join(parts[:-1]), "", parts[-1]) if len(parts) > 1 else (raw, "", "Barcelona")


def first_of(v):
    return re.split(r"\s+/\s*|\s*/\s+|\s+(?=https?://)", (v or "").strip())[0].strip()


def build(fields, urls, img, meta, base_dir):
    nome = meta.get("nome") or re.sub(r"\s+", " ", fields.get("nome", "")).strip()
    forma_raw = fields.get("forma", "")
    categoria = meta.get("categoria")
    if not categoria:
        for part in re.split(r"[/,]", forma_raw):
            n = norm(part)
            n = re.sub(r"^as ", "", n)
            if n in FORMA_MAP:
                categoria = FORMA_MAP[n]; break
    categoria = categoria or "associacao"
    areas = [a.strip(" .") for a in re.split(r"[,/;]", fields.get("area", "")) if a.strip(" .")]
    subcats = [AREA_MAP[norm(a)] for a in areas if norm(a) in AREA_MAP]

    media = {k: "" for k in ("facebook", "instagram", "linkedin", "twitter", "youtube", "tiktok")}
    excl = set(meta.get("excluirLinks", []))
    for u in urls:
        if u in excl:
            continue
        for dom, key in SOCIAL:
            if re.search(r"//(www\.|m\.)?" + re.escape(dom) + r"(/|$)", u) and not media.get(key):
                if key == "twitter":
                    u = re.sub(r"/status/.*$", "", u)
                media[key] = u
                break
    media = {k: v for k, v in media.items() if v or k in ("facebook", "instagram", "linkedin", "twitter", "youtube")}

    website = first_of(fields.get("website", ""))
    if website and not website.startswith("http"):
        website = "https://" + website
    email = first_of(fields.get("email", ""))
    tel = first_of(fields.get("telefone", ""))

    direccion = fields.get("direccion", "")
    street, cp, city = parse_address(direccion)
    c = meta.get("coords")
    full = ", ".join(x for x in [street, f"{cp} {city}".strip(), "España"] if x) if direccion else ""
    sobre = re.sub(r"^\d+\s+palabras\s+", "", fields.get("sobre", "").strip(), flags=re.I)
    rec = {
        "projeto": "es",
        "type": "fisica" if direccion else "digital",
        "nome": nome,
        "identificacao": "",
        "categoria": categoria,
        "formaJuridicaOriginal": forma_raw,
        "subCategoria": "; ".join(dict.fromkeys(subcats)),
        "areasOriginal": areas,
        "eMail": email,
        "telefone": tel,
        "website": website,
        "mediaLinks": media,
        "address": {
            "street": street if direccion else "", "number": "", "city": city if direccion else "",
            "country": "España", "postalCode": cp,
            "latitude": str(c[0]) if c and direccion else "",
            "longitude": str(c[1]) if c and direccion else "",
            "district": "", "municipality": city if direccion else "", "parish": "",
            "fullAddress": full,
        },
        "img": Path(img).relative_to(base_dir).as_posix() if img else "",
        "sobre": sobre,
        "sobre_i18n": {k: v for k, v in {
            "es": sobre,
            "pt": fields.get("sobre_pt", "").strip(),
            "en": fields.get("sobre_en", "").strip()}.items() if v},
    }
    extras = {"eMails": fields.get("email", ""), "telefones": fields.get("telefone", "")}
    for k, v in extras.items():
        if v and v != (email if k == "eMails" else tel):
            rec[k] = v
    # Correções manuais vindas do cliente (têm prioridade sobre a ficha)
    for k in ("telefone", "eMail", "eMails", "telefones", "website", "notaMorada"):
        if meta.get(k):
            rec[k] = meta[k]
    for lang, txt in (meta.get("sobre_i18n") or {}).items():
        rec["sobre_i18n"][lang] = txt
    if rec["sobre_i18n"].get("es"):
        rec["sobre"] = rec["sobre_i18n"]["es"]
    if meta.get("evento"):
        rec["evento"] = True
    if meta.get("coordsAproximadas") and c:
        rec["coordsAproximadas"] = True
    if fields.get("atividades"):
        rec["atividades"] = fields["atividades"]
    return rec


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", required=True)
    ap.add_argument("--meta", default=None)
    ap.add_argument("docx", nargs="+")
    a = ap.parse_args()

    data_path = Path(a.data)
    base = data_path.parent
    data = json.loads(data_path.read_text(encoding="utf-8"))
    metas = json.loads(Path(a.meta).read_text(encoding="utf-8")) if a.meta else {}

    by_name = {norm(d.get("nome")): i for i, d in enumerate(data) if d.get("projeto") == "es"}
    next_id = max([ID_START - 1] + [d.get("id", 0) for d in data if d.get("projeto") == "es"]) + 1
    warnings = 0
    for f in a.docx:
        fp = Path(f)
        meta = metas.get(fp.stem, {})
        if meta.get("ignorar"):
            nm = norm(meta.get("nome") or fp.stem)
            if nm in by_name:
                data.pop(by_name[nm])
                by_name = {norm(d.get("nome")): i for i, d in enumerate(data) if d.get("projeto") == "es"}
                print(f"✖ removido:  {fp.name}")
            else:
                print(f"· ignorado: {fp.name}")
            continue
        fields, urls = read_docx(f)
        if not fields.get("nome"):
            print(f"⚠ {fp.name}: sem NOMBRE — ignorado"); warnings += 1; continue
        name = meta.get("nome") or fields["nome"]
        logo = fp.parent / meta["logo"] if meta.get("logo") else None
        img = extract_image(f, base / "img" / "es", name, logo)
        rec = build(fields, urls, img, meta, base)
        key = norm(rec["nome"])
        old_key = norm(meta.get("nomeAnterior", "")) if meta.get("nomeAnterior") else None
        idx = by_name.get(key, by_name.get(old_key) if old_key else None)
        if idx is not None:
            old = data[idx]
            # mantém traduções já feitas (pt/en) que a ficha não traga
            for lang, txt in (old.get("sobre_i18n") or {}).items():
                rec["sobre_i18n"].setdefault(lang, txt)
            data[idx] = {"id": old["id"], "no": old["no"], **rec}
            by_name[key] = idx
            action = "atualizado"
        else:
            data.append({"id": next_id, "no": next_id, **rec})
            by_name[key] = len(data) - 1
            next_id += 1
            action = "novo"
        msg = f"✔ {action:10} {rec['nome']} [{rec['type']}/{rec['categoria']}] {rec['subCategoria'] or '(sem área c/ ícone)'}"
        if rec["type"] == "fisica" and not rec["address"]["latitude"]:
            msg += "  ⚠ sem coordenadas"; warnings += 1
        if not rec["img"]:
            msg += "  ⚠ sem logo"
        print(msg)

    data_path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"\n{len(data)} registos no total · {sum(1 for d in data if d.get('projeto') == 'es')} de Espanha · {warnings} aviso(s)")


if __name__ == "__main__":
    main()
