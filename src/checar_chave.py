# -*- coding: utf-8 -*-
r"""
CHECAGEM DA CHAVE — descobre se o problema é o arquivo ou a própria chave.
Não envia nada para lugar nenhum. Rode:  py checar_chave.py
"""
import pathlib, sys
HERE = pathlib.Path(__file__).parent
f = HERE / "api_key.txt"
if not f.exists():
    print("api_key.txt NÃO existe nesta pasta."); sys.exit(1)

bruto = f.read_bytes()
txt   = bruto.decode("utf-8", "ignore")
linhas = [l for l in txt.splitlines() if l.strip()]

print("="*66); print("CHECAGEM DO api_key.txt"); print("="*66)
print(f"  tamanho do arquivo ....... {len(bruto)} bytes")
print(f"  linhas com conteúdo ...... {len(linhas)}")
BOM = b"\xef\xbb\xbf"
tem_bom = "SIM — isso quebra a chave" if bruto[:3]==BOM else "não"
print(f"  começa com BOM ........... {tem_bom}")
if not linhas: print("\n  ARQUIVO VAZIO."); sys.exit(1)

k = linhas[0]
lim = k.strip()
tem_esp = "SIM — o script já remove" if k!=lim else "não"
print(f"  espaços nas pontas ....... {tem_esp}")
tem_aspas = "SIM — APAGUE as aspas" if (lim[:1] == chr(34) or lim[:1] == chr(39)) else "não"
print(f"  aspas em volta ........... {tem_aspas}")
print(f"  comprimento da chave ..... {len(lim)} caracteres")
print(f"  início ................... {lim[:20]}")
print(f"  fim ...................... ...{lim[-8:]}")
if "COLE A CHAVE" in txt.upper():
    print("\n  >> O arquivo ainda tem o texto de exemplo. Apague e cole a chave.")
if len(linhas) > 1:
    print(f"\n  >> Há {len(linhas)} linhas. Deixe SÓ a chave, em uma linha.")

print("\n" + "="*66); print("QUAL API ESSA CHAVE ABRE?"); print("="*66)
if lim.startswith("iclips_sk_"):
    print("  Formato: iclips_sk_...")
    print("  Esta é a chave da API DE PROJETOS (a mesma usada no Power BI).")
    print("  A Public API (public-api.iclips.com.br) espera OUTRA chave, gerada em")
    print("  iClips > Avatar > CHAVE DE API, cujo formato começa com 'eyJ...'.")
    print("\n  >> Provável causa do 401: chave certa, API errada.")
elif lim.startswith("eyJ"):
    print("  Formato: eyJ... — é o formato esperado pela Public API.")
    print("  Se mesmo assim deu 401: a chave pode estar incompleta (confira se copiou")
    print("  tudo) ou a agência não está no Plano PRO.")
else:
    print(f"  Formato não reconhecido (começa com '{lim[:6]}').")
    print("  A Public API espera uma chave começando com 'eyJ'.")
print("="*66)
