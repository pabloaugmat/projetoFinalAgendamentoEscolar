import pandas as pd
import igraph as ig
import matplotlib.pyplot as plt
import math
import numpy as np

def ler_excel(nome_arquivo, skip_rows, num_rows):
    """Carrega um arquivo Excel em um DataFrame."""
    df = pd.read_excel(nome_arquivo, skiprows=skip_rows, nrows=num_rows)
    return df

def adcionarTAG(df, ccoRange, sinRange, optRange, posRange):
    disciplinas = df["SIGLA"].tolist()  # Converter a coluna "SIGLA" do DataFrame em uma lista de disciplinas
    disciplinasCCO = [disciplina + "[CCO]" for disciplina in disciplinas[:ccoRange]]
    disciplinasSIN = [disciplina + "[SIN]" for disciplina in disciplinas[ccoRange:ccoRange+sinRange]]
    disciplinasOPT = [disciplina + "[OPT]" for disciplina in disciplinas[ccoRange+sinRange:ccoRange+sinRange+optRange]]
    disciplinasPOS = [disciplina + "[POS]" for disciplina in disciplinas[ccoRange+sinRange+optRange:ccoRange+sinRange+optRange+posRange]]
    disciplinasTAG = disciplinasCCO+disciplinasSIN+disciplinasOPT+disciplinasPOS
    df["SIGLA"] = disciplinasTAG
        
def gerar_grafo(df, ccoRange, sinRange, optRange, posRange):
    #remoção de colunas dispensaveis para geração do grafo
    df.drop(["Unnamed: 0", "Unnamed: 1", "PPC", "DISCIPLINA", "CH"], axis=1, inplace=True)
    g = ig.Graph()  # Criar um novo grafo
    ### VÉRTICES ###
    disciplinasTAG = df["SIGLA"].tolist()
    g.add_vertices(disciplinasTAG)  # Adicionar vértices ao grafo com as disciplinas
    # Definir cores para os vértices de acordo com as faixas especificadas
    vcores = (['red'] * ccoRange) + (['blue'] * sinRange) + (['green'] * optRange) + (['yellow'] * posRange)
    g.vs["color"] = vcores  # Atribuir cores aos vértices do grafo
    # Arestas entre vértices do mesmo período e mesma cor
    for faixa in [range(ccoRange), range(ccoRange, ccoRange+sinRange)]:
        for i in faixa:
            for j in faixa:
                if i < j:
                    disciplinaX = disciplinasTAG[i]
                    disciplinaY = disciplinasTAG[j]
                    idxX = g.vs.find(disciplinaX).index
                    idxY = g.vs.find(disciplinaY).index
                    if df.loc[df["SIGLA"] == disciplinaX, "PER."].values[0] == df.loc[df["SIGLA"] == disciplinaY, "PER."].values[0] and g.vs[idxX]["color"] == g.vs[idxY]["color"]:
                        g.add_edge(disciplinaX, disciplinaY, color='magenta')
    # arestas de professores
    # Criar arestas baseadas em valores numéricos nas colunas de professores
    professores = [col for col in df.columns if 'prof' in col]  # Obter lista de colunas dos professores
    for professor in professores:
        indices_validos = []
        for idx, valor in df[professor].items():
            if not math.isnan(valor):  # Verificar se o valor não é NaN
                indices_validos.append(idx)
        # Criar arestas entre todos os pares de índices válidos
        for i in range(len(indices_validos)):
            for j in range(i + 1, len(indices_validos)):
                idxX = indices_validos[i]
                idxY = indices_validos[j]
                disciplinaX = disciplinasTAG[idxX]
                disciplinaY = disciplinasTAG[idxY]
                g.add_edge(disciplinaX, disciplinaY, color='cyan')        
    return g

def plotar_e_salvar_grafo(g, nome_arquivo):
    """Plota e salva o grafo usando igraph."""
    layout = g.layout("fr")  # Fruchterman-Reingold layout
    visual_style = {
        "vertex_label": g.vs["name"],
        "vertex_size": 20,
        "bbox": (1600, 1200),
        "margin": 50,
        "layout": layout
    }
    # Salvar o grafo em um arquivo
    ig.plot(g, **visual_style).save(nome_arquivo)
    # Plotar o grafo
    ig.plot(g, **visual_style)
    plt.show()
    
def gerar_matriz_simples(grafo):
    return grafo.get_adjacency()
    
def gerar_matriz_dataframe(grafo):
    matriz = gerar_matriz_simples(grafo)
    dataFrame = pd.DataFrame(matriz.data, columns=grafo.vs["name"], index=grafo.vs["name"])
    print(dataFrame)
    return dataFrame

def gerar_complemento(grafo):
    complemento = grafo.complementer()
    return complemento

def dict_aulaCH(df, ccoRange, sinRange, optRange, posRange):
    dict_sigla_ch = df.set_index('SIGLA')['CH'].to_dict()
    return dict_sigla_ch

def get_key_index(dictionary, key):
    keys_list = list(dictionary.keys())
    return keys_list.index(key)

def conflito(matrix_adj, idx1, idx2):
    if matrix_adj[idx1][idx2] > 0:
        return True
    else:
        return False

def gerar_tabela(grafo, aulasCH, nome_arquivo):
    data = {
        'SEGUNDA': [[] for _ in range(15)],
        'TERCA': [[] for _ in range(15)],
        'QUARTA': [[] for _ in range(15)],
        'QUINTA': [[] for _ in range(15)],
        'SEXTA': [[] for _ in range(15)]
    }
    # Criando o DataFrame
    tabela = pd.DataFrame(data)
    # Definindo os rótulos para as linhas
    horarios = ["M1", "M2", "M3", "M4", "M5", "T1", "T2", "T3", "T4", "T5", "N1", "N2", "N3", "N4", "N5"]
    tabela.index = horarios
    dias = tabela.columns.values
    matriz_adj = gerar_matriz_simples(grafo)
    aulas = list(aulasCH.keys())
    aulasTotal = []
    for item in aulas:
        for i in range(0, aulaCH[item]):
            aulasTotal.append(item)
    # ALOCAÇÃO DOS AULAS EM HORARIOS PELOS DIAS DA SEMANA EVITANDO CONFLITOS
    distribuicaoDiaria = len(aulasTotal)/5
    for dia in dias:
        aulas_no_dia = []
        for horario in horarios:
            if len(aulas_no_dia) < distribuicaoDiaria:
                for aula in aulasTotal:
                    haConflito = False
                    if aula not in tabela.at[horario, dia]:
                        for item in tabela.at[horario, dia]:
                            if conflito(matriz_adj, get_key_index(aulaCH, aula), get_key_index(aulaCH, item)):
                                haConflito = True
                        if not haConflito:
                            if aulas_no_dia.count(aula) < 2:
                                aulas_no_dia.append(aula)
                                tabela.at[horario, dia].append(aula)
                                aulasTotal.remove(aula) 
    tabela.to_excel(nome_arquivo, index=True)

if __name__ == "__main__":
    # DATAFRAME
    nome_arquivo = "../data/cenario6_dados.xlsx"
    # GRAFO 2024_1
    df = ler_excel(nome_arquivo, 7, 49)
    ccoRange = 21
    sinRange = 19
    optRange = 7
    posRange = 1
    adcionarTAG(df, ccoRange, sinRange, optRange, posRange)
    aulaCH = dict_aulaCH(df, ccoRange, sinRange, optRange, posRange)
    g2024_1 = gerar_grafo(df, ccoRange, sinRange, optRange, posRange)
    plotar_e_salvar_grafo(g2024_1, "../data/grafo_disciplinas_2024_1.png")
    # grafo complemento 2024_1
    complemento = gerar_complemento(g2024_1)
    plotar_e_salvar_grafo(complemento, "../data/complemento_disciplinas_2024_1.png")
    # Gera tabela xlsx com as aulas alocadas em horarios para o 2024_1
    gerar_tabela(g2024_1, aulaCH, "../data/agenda_aulas_2024_1.xlsx")
    # GRAFO 2024_2
    df = ler_excel(nome_arquivo, 59, 41)
    ccoRange = 13
    sinRange = 18
    optRange = 7
    posRange = 2
    adcionarTAG(df, ccoRange, sinRange, optRange, posRange)
    aulaCH = dict_aulaCH(df, ccoRange, sinRange, optRange, posRange)
    g2024_2 = gerar_grafo(df, ccoRange, sinRange, optRange, posRange)
    plotar_e_salvar_grafo(g2024_2, "../data/grafo_disciplinas_2024_2.png")
    # grafo complemento 2024_2
    complemento = gerar_complemento(g2024_2)
    plotar_e_salvar_grafo(complemento, "../data/complemento_disciplinas_2024_2.png")
    # Gera tabela xlsx com as aulas alocadas em horarios para o 2024_2
    gerar_tabela(g2024_2, aulaCH, "../data/agenda_aulas_2024_2.xlsx")

    