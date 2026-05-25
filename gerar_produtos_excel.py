# ─────────────────────────────────────────────────────────────────────────────
# gerar_produtos_excel.py
#
# Execute UMA VEZ no terminal para criar o arquivo produtos.xlsx:
#   python gerar_produtos_excel.py
#
# Depois suba o "produtos.xlsx" para o GitHub junto com o app.py.
# Para atualizar preços: edite o Excel e faça commit — sem mexer no código!
# ─────────────────────────────────────────────────────────────────────────────

import pandas as pd

PRODUTOS = [
    {"area":"Interna","colecao":"Taquara","movel":"Mesa Centro","cor":"amber","bp":"Antares coffee table","desc":"mesa de centro","preco":3870,"custo":851.28},
    {"area":"Interna","colecao":"Taquara","movel":"Aparador","cor":"amber","bp":"Antares console table","desc":"aparador","preco":3110,"custo":596.47},
    {"area":"Interna","colecao":"Taquara","movel":"Mesa lateral","cor":"amber","bp":"Antares end table","desc":"mesa lateral","preco":1990,"custo":357.90},
    {"area":"Interna","colecao":"Guadua","movel":"Banco","cor":"caramelized","bp":"azara bench","desc":"banco caramelo","preco":2970,"custo":441.93},
    {"area":"Interna","colecao":"Guadua","movel":"Banco","cor":"sable","bp":"azara bench","desc":"banco café","preco":2970,"custo":441.93},
    {"area":"Interna","colecao":"Alyra","movel":"Poltrona","cor":"caramelized/red","bp":"Danica chair","desc":"poltrona caramelo/vermelho","preco":3540,"custo":577.95},
    {"area":"Interna","colecao":"Alyra","movel":"Poltrona","cor":"caramelized/grey","bp":"Danica chair","desc":"poltrona caramelo/cinza","preco":3540,"custo":425.46},
    {"area":"Interna","colecao":"Balcoa","movel":"Mesa de centro","cor":"amber","bp":"rosemary coffee table","desc":"mesa de centro","preco":4180,"custo":646.51},
    {"area":"Interna","colecao":"Balcoa","movel":"Mesa de centro","cor":"wheat","bp":"rosemary coffee table","desc":"mesa de centro","preco":4180,"custo":516.08},
    {"area":"Interna","colecao":"Balcoa","movel":"Mesa lateral","cor":"amber","bp":"thyme side table","desc":"mesa lateral","preco":1490,"custo":385.35},
    {"area":"Interna","colecao":"Balcoa","movel":"Mesa lateral","cor":"wheat","bp":"thyme side table","desc":"mesa lateral","preco":1490,"custo":363.51},
    {"area":"Interna","colecao":"Aurea","movel":"Mesa lateral","cor":"amber","bp":"sol side table","desc":"mesa lateral redonda","preco":1780,"custo":451.05},
    {"area":"Interna","colecao":"Aurea","movel":"Mesa lateral","cor":"wheat","bp":"sol side table","desc":"mesa lateral redonda","preco":1780,"custo":451.05},
    {"area":"Interna","colecao":"Aurea","movel":"Mesa de cabeceira","cor":"caramelized","bp":"currant nightstand","desc":"mesa de cabeceira","preco":3650,"custo":789.76},
    {"area":"Interna","colecao":"Aurea","movel":"Mesa de cabeceira","cor":"amber","bp":"currant nightstand","desc":"mesa de cabeceira","preco":3650,"custo":789.76},
    {"area":"Interna","colecao":"Aurea","movel":"Comoda","cor":"caramelized","bp":"currant dresser","desc":"comoda","preco":13460,"custo":3245.65},
    {"area":"Interna","colecao":"Aurea","movel":"Comoda","cor":"amber","bp":"currant dresser","desc":"comoda","preco":13890,"custo":3245.65},
    {"area":"Interna","colecao":"Aurea","movel":"Estante","cor":"caramelized","bp":"currant leaning bookshelf","desc":"estante de parede","preco":3590,"custo":541.14},
    {"area":"Interna","colecao":"Aurea","movel":"Estante","cor":"black walnut","bp":"currant leaning bookshelf","desc":"estante de parede","preco":3790,"custo":541.14},
    {"area":"Interna","colecao":"Kuruna","movel":"Mesa de cabeceira","cor":"wheat","bp":"monterey nightstand","desc":"mesa de cabeceira","preco":3290,"custo":813.28},
    {"area":"Interna","colecao":"Kuruna","movel":"Comoda","cor":"wheat","bp":"monterey dresser","desc":"comoda","preco":12220,"custo":3323.54},
    {"area":"Interna","colecao":"Valiha","movel":"Mesa de cabeceira","cor":"amber","bp":"ventura nightstand","desc":"mesa de cabeceira","preco":3305,"custo":813.28},
    {"area":"Interna","colecao":"Valiha","movel":"Comoda","cor":"amber","bp":"ventura dresser","desc":"comoda","preco":12220,"custo":3323.53},
    {"area":"Interna","colecao":"Arberella","movel":"Mesa de cabeceira","cor":"caramelized","bp":"willow nightstand","desc":"mesa de cabeceira","preco":3170,"custo":804.17},
    {"area":"Interna","colecao":"Arberella","movel":"Comoda","cor":"caramelized","bp":"willow dresser","desc":"comoda tradicional","preco":9110,"custo":2478.02},
    {"area":"Interna","colecao":"Sasa","movel":"Mesa de jantar","cor":"amber (erikka)","bp":"erikka dining table","desc":"mesa jantar variável","preco":14890,"custo":3399.10},
    {"area":"Interna","colecao":"Sasa","movel":"Mesa de jantar","cor":"amber (mija)","bp":"mija dining table","desc":"mesa jantar menor","preco":9980,"custo":1094.56},
    {"area":"Interna","colecao":"Sasa","movel":"Cadeira","cor":"amber","bp":"currant chair","desc":"cadeira clean","preco":1980,"custo":534.44},
    {"area":"Interna","colecao":"Sasa","movel":"Banqueta","cor":"caramelized","bp":"tulip bar stoll","desc":"banqueta curva","preco":1780,"custo":509.52},
    {"area":"Interna","colecao":"Pinga","movel":"Mesa de jantar","cor":"wheat","bp":"hanna dining table","desc":"mesa jantar orgânica 6 lug","preco":15990,"custo":1904.30},
    {"area":"Interna","colecao":"Pinga","movel":"Cadeira","cor":"wheat","bp":"hanna dining chair","desc":"cadeira design","preco":2680,"custo":697.57},
    {"area":"Interna","colecao":"Pinga","movel":"Aparador","cor":"wheat","bp":"hanna console/sideboard","desc":"aparador bambu","preco":12260,"custo":1126.52},
    {"area":"Interna","colecao":"Tulda","movel":"Mesa","cor":"wheat","bp":"sitka dining table","desc":"redonda 4 lugares","preco":4990,"custo":601.33},
    {"area":"Interna","colecao":"Tulda","movel":"Mesa","cor":"amber","bp":"sitka dining table","desc":"redonda 4 lugares","preco":5110,"custo":3106.62},
    {"area":"Interna","colecao":"Pariana","movel":"Banqueta","cor":"wheat","bp":"leif counter stoll","desc":"banqueta corda caramelo","preco":2110,"custo":645.38},
    {"area":"Interna","colecao":"Pariana","movel":"Banqueta","cor":"caviar","bp":"leif counter stoll","desc":"banqueta corda preta","preco":2110,"custo":645.38},
    {"area":"Externa","colecao":"Taquara","movel":"Mesa","cor":"—","bp":"Bled Table 210cm","desc":"mesa externa","preco":8110,"custo":895.29},
    {"area":"Externa","colecao":"Taquara","movel":"Cadeira","cor":"—","bp":"Bled Armchair","desc":"cadeira externa","preco":1940,"custo":527.40},
    {"area":"Externa","colecao":"Taquara","movel":"Espreguiçadeira","cor":"—","bp":"Bled Sun Lounger","desc":"espreguiçadeira","preco":5990,"custo":820.72},
    {"area":"Externa","colecao":"Taquara","movel":"Ombrelone","cor":"—","bp":"AA Sun Umbrella","desc":"ombrelone","preco":6530,"custo":1699.85},
    {"area":"Externa","colecao":"Pariana","movel":"Mesa","cor":"—","bp":"Positano Table 215cm","desc":"mesa externa","preco":8230,"custo":973.44},
    {"area":"Externa","colecao":"Pariana","movel":"Cadeira","cor":"—","bp":"Positano Chair","desc":"cadeira externa","preco":2600,"custo":742.54},
    {"area":"Externa","colecao":"Balcoa","movel":"Mesa redonda","cor":"—","bp":"Sunflower Table 98cm","desc":"mesa redonda","preco":5230,"custo":1147.60},
    {"area":"Externa","colecao":"Balcoa","movel":"Espreguiçadeira","cor":"—","bp":"Flexi Sun Lounger","desc":"espreguiçadeira","preco":11525,"custo":841.95},
    {"area":"Externa","colecao":"Alvimia","movel":"Sofá 3 lugares","cor":"—","bp":"Koloni 3-seater Sofa","desc":"sofá 3 lugares","preco":13000,"custo":2282.08},
    {"area":"Externa","colecao":"Alvimia","movel":"Sofá 1 lugar","cor":"—","bp":"Koloni Sofa","desc":"sofá 1 lugar","preco":5250,"custo":292.12},
    {"area":"Externa","colecao":"Alvimia","movel":"Mesa de centro","cor":"—","bp":"Koloni Table","desc":"mesa de centro","preco":5270,"custo":887.08},
    {"area":"Externa","colecao":"Nadinha","movel":"Cadeira","cor":"—","bp":"ET Chair","desc":"cadeira","preco":2650,"custo":689.46},
    {"area":"Externa","colecao":"Nadinha","movel":"Mesa","cor":"—","bp":"ET Table","desc":"mesa","preco":3120,"custo":695.83},
]

df = pd.DataFrame(PRODUTOS)
df.to_excel("produtos.xlsx", index=False)
print(f"✅ produtos.xlsx criado com {len(df)} produtos!")
print("Agora suba este arquivo para o GitHub junto com o app.py.")
