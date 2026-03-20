"""
Execute UMA VEZ para regerar os .pbids com o formato correto.
Pode deletar depois.
"""
import json, glob, os

PASTA = r'Z:\GUSTAVO\App\Consolidados'

MODULOS_MAP = {
    'Pedidos_e_Recebimentos': 'PowerBI_Pedidos',
    'Fretes':                 'PowerBI_Fretes',
    'Armazenagem':            'PowerBI_Armazenagem',
    'Estoque':                'PowerBI_Estoque',
    'Produtividade':          'PowerBI_Produtividade',
    'consolidado':            'PowerBI_Dashboard',
}

gerados = []
for prefixo, nome_pbids in MODULOS_MAP.items():
    matches = sorted(
        glob.glob(os.path.join(PASTA, f'{prefixo}*.xlsx')),
        key=os.path.getmtime, reverse=True
    )
    if not matches:
        print(f'⚠️  {prefixo}: nenhum .xlsx encontrado — pulando')
        continue

    caminho_xlsx = os.path.abspath(matches[0]).replace('/', '\\')
    pbids = {
        "version": "0.1",
        "connections": [{
            "details": {
                "protocol": "file",
                "address": { "path": caminho_xlsx }
            },
            "mode": "Import"
        }]
    }

    out = os.path.join(PASTA, f'{nome_pbids}.pbids')
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(pbids, f, ensure_ascii=False, indent=2)
    gerados.append(f'{nome_pbids}.pbids → {os.path.basename(caminho_xlsx)}')
    print(f'✅ {nome_pbids}.pbids')

print(f'\n{len(gerados)} arquivo(s) gerados em:\n{PASTA}')
input('\nPressione Enter para fechar...')