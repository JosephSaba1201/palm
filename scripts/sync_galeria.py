#!/usr/bin/env python3
"""
Sincroniza la carpeta fotos/ con el arreglo GALERIA_RAW embebido en index.html.

Que hace en cada corrida:
  1. Lee todas las fotos en fotos/ (excluye thumbs/ y archivos ocultos).
  2. Para las que ya estan en GALERIA_RAW (por nombre de archivo), conserva su
     fecha/categoria/descripcion tal cual.
  3. Para las que son nuevas, determina la fecha real en este orden:
       a) fecha EXIF (DateTimeOriginal) si el archivo la trae,
       b) fecha en el nombre del archivo (WhatsApp Image..., PHOTO-...,
          o el propio patron NN_YYYY-MM-DD),
       c) si no se puede determinar con certeza, NO se adivina: se usa la
          fecha de modificacion del archivo y se marca con
          categoria "Revisar" para que quede visible que hay que confirmarla.
  4. Renumera TODO en orden cronologico (01_, 02_, ...) y renombra los
     archivos fisicos en fotos/ y fotos/thumbs/ para que coincidan.
  5. Genera con `sips` la miniatura (700px de ancho) de cualquier foto nueva
     que no la tenga todavia.
  6. Hace un backup de index.html en .backups/ y reescribe por completo el
     bloque GALERIA_RAW con el resultado final.

Uso:
  python3 scripts/sync_galeria.py            # aplica los cambios
  python3 scripts/sync_galeria.py --dry-run  # solo muestra que haria, no toca nada
"""
import os, re, sys, shutil, subprocess, struct
from datetime import datetime

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FOTOS = os.path.join(BASE, "fotos")
THUMBS = os.path.join(FOTOS, "thumbs")
HTML = os.path.join(BASE, "index.html")
BACKUPS = os.path.join(BASE, ".backups")

NUM_DATE_RE = re.compile(r'^\d{2}_(\d{4}-\d{2}-\d{2})\.[A-Za-z]+$')
WHATSAPP_RE = re.compile(r'(\d{4}-\d{2}-\d{2})')
IMG_EXT = {'.jpg', '.jpeg', '.png', '.JPG', '.JPEG', '.PNG'}


def read_exif_date(path):
    """Lee DateTimeOriginal (0x9003) o DateTime (0x0132) de un JPEG. Solo stdlib."""
    try:
        with open(path, 'rb') as f:
            if f.read(2) != b'\xff\xd8':
                return None
            while True:
                marker = f.read(2)
                if len(marker) < 2 or marker[0] != 0xFF:
                    return None
                if marker[1] == 0xE1:  # APP1
                    size = struct.unpack('>H', f.read(2))[0]
                    data = f.read(size - 2)
                    if not data.startswith(b'Exif\x00\x00'):
                        continue
                    tiff = data[6:]
                    if tiff[:2] == b'II':
                        endian = '<'
                    elif tiff[:2] == b'MM':
                        endian = '>'
                    else:
                        return None
                    ifd0_off = struct.unpack(endian + 'I', tiff[4:8])[0]

                    def read_ifd(offset):
                        n = struct.unpack(endian + 'H', tiff[offset:offset+2])[0]
                        entries = {}
                        for i in range(n):
                            e = tiff[offset+2+i*12: offset+2+i*12+12]
                            tag, typ, count = struct.unpack(endian + 'HHI', e[:8])
                            val_raw = e[8:12]
                            entries[tag] = (typ, count, val_raw)
                        return entries

                    ifd0 = read_ifd(ifd0_off)
                    for tag in (0x8769,):  # Exif SubIFD pointer
                        if tag in ifd0:
                            sub_off = struct.unpack(endian + 'I', ifd0[tag][2])[0]
                            sub = read_ifd(sub_off)
                            for dtag in (0x9003, 0x9004):
                                if dtag in sub:
                                    _, count, val_raw = sub[dtag]
                                    off = struct.unpack(endian + 'I', val_raw)[0]
                                    raw = tiff[off:off+count].rstrip(b'\x00')
                                    s = raw.decode('ascii', 'ignore')
                                    dt = datetime.strptime(s, '%Y:%m:%d %H:%M:%S')
                                    return dt.strftime('%Y-%m-%d')
                    if 0x0132 in ifd0:
                        _, count, val_raw = ifd0[0x0132]
                        if count <= 4:
                            raw = val_raw[:count].rstrip(b'\x00')
                        else:
                            off = struct.unpack(endian + 'I', val_raw)[0]
                            raw = tiff[off:off+count].rstrip(b'\x00')
                        s = raw.decode('ascii', 'ignore')
                        dt = datetime.strptime(s, '%Y:%m:%d %H:%M:%S')
                        return dt.strftime('%Y-%m-%d')
                    return None
                elif marker[1] in (0xD8, 0x01) or 0xD0 <= marker[1] <= 0xD9:
                    continue
                else:
                    size = struct.unpack('>H', f.read(2))[0]
                    f.seek(size - 2, 1)
    except Exception:
        return None
    return None


def guess_date(fname, fpath):
    m = NUM_DATE_RE.match(fname)
    if m:
        return m.group(1), 'nombre (ya numerado)'
    exif = read_exif_date(fpath)
    if exif:
        return exif, 'EXIF'
    m = WHATSAPP_RE.search(fname)
    if m:
        return m.group(1), 'nombre de archivo'
    mtime = datetime.fromtimestamp(os.path.getmtime(fpath)).strftime('%Y-%m-%d')
    return mtime, 'fecha de modificacion (sin certeza)'


def parse_galeria_raw(html):
    m = re.search(r'const GALERIA_RAW = \[(.*?)\n\];', html, re.S)
    if not m:
        raise SystemExit("No se encontro GALERIA_RAW en index.html")
    block = m.group(1)
    entries = []
    for line in block.split('\n'):
        line = line.strip().rstrip(',')
        if not line.startswith('{file:'):
            continue
        fm = re.search(r'file:"([^"]*)"', line)
        im = re.search(r'iso:"([^"]*)"', line)
        cm = re.search(r'categoria:"([^"]*)"', line)
        capm = re.search(r'cap:"((?:[^"\\]|\\.)*)"', line)
        entries.append({
            'file': fm.group(1),
            'iso': im.group(1),
            'categoria': cm.group(1),
            'cap': capm.group(1) if capm else '',
        })
    return entries, m.span(1)


def main():
    dry = '--dry-run' in sys.argv

    with open(HTML, encoding='utf-8') as f:
        html = f.read()
    existing, _ = parse_galeria_raw(html)
    existing_by_file = {e['file']: e for e in existing}

    disk_files = sorted(
        f for f in os.listdir(FOTOS)
        if os.path.isfile(os.path.join(FOTOS, f))
        and not f.startswith('.')
        and os.path.splitext(f)[1] in IMG_EXT
    )

    missing_on_disk = [e['file'] for e in existing if e['file'] not in disk_files]
    if missing_on_disk:
        print("AVISO: estas fotos estan en el codigo pero ya no existen en fotos/ (no se tocan, revisar a mano):")
        for f in missing_on_disk:
            print(f"   - {f}")

    records = []
    new_flagged = []
    for fname in disk_files:
        fpath = os.path.join(FOTOS, fname)
        if fname in existing_by_file:
            e = existing_by_file[fname]
            records.append({'old_name': fname, 'iso': e['iso'], 'categoria': e['categoria'], 'cap': e['cap']})
        else:
            iso, source = guess_date(fname, fpath)
            categoria = 'General' if source != 'fecha de modificacion (sin certeza)' else 'Revisar'
            cap = 'Pendiente de descripcion.' if source != 'fecha de modificacion (sin certeza)' \
                else f'Fecha sin confirmar (tomada de {source}) — revisar y corregir categoria/fecha.'
            records.append({'old_name': fname, 'iso': iso, 'categoria': categoria, 'cap': cap})
            new_flagged.append((fname, iso, source))

    if new_flagged:
        print(f"Fotos nuevas detectadas ({len(new_flagged)}):")
        for fname, iso, source in new_flagged:
            print(f"   + {fname}  ->  iso={iso}  (fuente: {source})")

    records.sort(key=lambda r: (r['iso'], r['old_name']))

    renames = []
    for i, r in enumerate(records, start=1):
        ext = os.path.splitext(r['old_name'])[1]
        new_name = f"{i:02d}_{r['iso']}{ext}"
        r['new_name'] = new_name
        if new_name != r['old_name']:
            renames.append((r['old_name'], new_name))

    print(f"\nTotal fotos: {len(records)}. Renombres necesarios: {len(renames)}.")
    for old, new in renames:
        print(f"   {old}  ->  {new}")

    if dry:
        print("\n(--dry-run) no se modifico nada.")
        return

    new_names = [r['new_name'] for r in records]
    if len(new_names) != len(set(new_names)):
        raise SystemExit("Colision de nombres al renumerar. Abortando sin tocar nada.")

    # renombrar usando nombres temporales primero para evitar colisiones cruzadas
    tmp_suffix = '.tmp_sync'
    for old, new in renames:
        os.rename(os.path.join(FOTOS, old), os.path.join(FOTOS, old + tmp_suffix))
        thumb_old = os.path.join(THUMBS, old)
        if os.path.isfile(thumb_old):
            os.rename(thumb_old, thumb_old + tmp_suffix)
    for old, new in renames:
        os.rename(os.path.join(FOTOS, old + tmp_suffix), os.path.join(FOTOS, new))
        thumb_old_tmp = os.path.join(THUMBS, old + tmp_suffix)
        if os.path.isfile(thumb_old_tmp):
            os.rename(thumb_old_tmp, os.path.join(THUMBS, new))

    os.makedirs(THUMBS, exist_ok=True)
    for r in records:
        thumb_path = os.path.join(THUMBS, r['new_name'])
        if not os.path.isfile(thumb_path):
            src = os.path.join(FOTOS, r['new_name'])
            subprocess.run(['sips', '--resampleWidth', '700', src, '--out', thumb_path],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            print(f"   thumbnail generado: {r['new_name']}")

    os.makedirs(BACKUPS, exist_ok=True)
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    shutil.copy(HTML, os.path.join(BACKUPS, f'index_pre_sync_galeria_{stamp}.html'))

    def esc(s):
        return s.replace('\\', '\\\\').replace('"', '\\"')

    lines = []
    for r in records:
        lines.append(
            f'  {{file:"{esc(r["new_name"])}", iso:"{r["iso"]}", '
            f'categoria:"{esc(r["categoria"])}", cap:"{esc(r["cap"])}"}},'
        )
    new_block = '\n' + '\n'.join(lines) + '\n'

    with open(HTML, encoding='utf-8') as f:
        html = f.read()
    m = re.search(r'(const GALERIA_RAW = \[)(.*?)(\n\];)', html, re.S)
    html = html[:m.start(2)] + new_block + html[m.end(2):]
    with open(HTML, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"\nListo. index.html actualizado ({len(records)} fotos en GALERIA_RAW).")
    if new_flagged:
        print("Revisa las fotos marcadas como nuevas: confirma fecha, categoria y escribe una descripcion real.")


if __name__ == '__main__':
    main()
