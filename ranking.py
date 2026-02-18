import requests
import pandas as pd
import os
import datetime

# --- CONFIGURAÇÃO (O GitHub vai ler isso das 'Secrets') ---
CLIENT_ID = os.environ.get('CLIENT_ID')
CLIENT_SECRET = os.environ.get('CLIENT_SECRET')
REFRESH_TOKEN = os.environ.get('REFRESH_TOKEN')
CLUB_ID = '1909864' # Coloque o número do seu clube aqui
NOME_ARQUIVO = 'Ranking_CNB_500km_2026.xlsx'

def obter_access_token():
    payload = {
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'refresh_token': REFRESH_TOKEN,
        'grant_type': 'refresh_token'
    }
    res = requests.post("https://www.strava.com/oauth/token", data=payload).json()
    return res.get('access_token')

def formatar_km(valor):
    return f"{valor:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".") + " km"

def formatar_alt(valor):
    return f"{int(valor):,}".replace(",", ".") + " m"

# 1. Tentar carregar planilha existente no repositório
if os.path.exists(NOME_ARQUIVO):
    with pd.ExcelFile(NOME_ARQUIVO) as reader:
        df_ranking = pd.read_excel(reader, sheet_name='Ranking')
        # Limpa formatação para cálculo
        if df_ranking['KM Total'].dtype == object:
            df_ranking['KM Total'] = df_ranking['KM Total'].str.replace(' km', '').str.replace('.', '').str.replace(',', '.').astype(float)
        if df_ranking['Altimetria (m)'].dtype == object:
            df_ranking['Altimetria (m)'].str.replace(' m', '').str.replace('.', '').astype(float)
        df_ranking = df_ranking.set_index('Atleta')
        
        df_historico = pd.read_excel(reader, sheet_name='IDs_Processados')
        ids_ja_somados = set(df_historico['id'].astype(str).tolist())
else:
    df_ranking = pd.DataFrame(columns=['Atleta', 'KM Total', 'Altimetria (m)']).set_index('Atleta')
    ids_ja_somados = set()

# 2. Puxar dados do Strava
access_token = obter_access_token()
if access_token:
    for pagina in range(1, 6):
        url = f"https://www.strava.com/api/v3/clubs/{CLUB_ID}/activities"
        atividades = requests.get(url, headers={'Authorization': f'Bearer {access_token}'}, 
                                  params={'per_page': 200, 'page': pagina}).json()
        if not atividades or 'errors' in atividades or len(atividades) == 0: break

        for act in atividades:
            id_unico = f"{act.get('distance')}_{act.get('elapsed_time')}_{act.get('athlete', {}).get('lastname')}"
            if id_unico not in ids_ja_somados:
                nome = f"{act.get('athlete', {}).get('firstname', 'Atleta')} {act.get('athlete', {}).get('lastname', '')}".strip()
                dist_km = act.get('distance', 0) / 1000
                alt = act.get('total_elevation_gain', 0)
                if dist_km > 0:
                    if nome not in df_ranking.index: df_ranking.loc[nome] = [0.0, 0.0]
                    df_ranking.at[nome, 'KM Total'] += dist_km
                    df_ranking.at[nome, 'Altimetria (m)'] += alt
                    ids_ja_somados.add(id_unico)

    # 3. Ordenar e Formatar
    df_ranking = df_ranking.sort_values(by='KM Total', ascending=False).reset_index()
    df_visual = df_ranking.copy()
    df_visual['KM Total'] = df_visual['KM Total'].apply(formatar_km)
    df_visual['Altimetria (m)'] = df_visual['Altimetria (m)'].apply(formatar_alt)

    # 4. Salvar localmente (o GitHub fará o 'upload' depois)
    with pd.ExcelWriter(NOME_ARQUIVO) as writer:
        df_visual.to_excel(writer, sheet_name='Ranking', index=False)
        pd.DataFrame(list(ids_ja_somados), columns=['id']).to_excel(writer, sheet_name='IDs_Processados', index=False)