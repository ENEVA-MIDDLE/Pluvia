from dotenv import load_dotenv
from calendar import monthrange
from dateutil.relativedelta import relativedelta as delta, FR
from dateutil.rrule import rrule, DAILY
from datetime import date
import pandas as pd
import numpy as np
import os 

load_dotenv()

from Postgres.Postgres import Postgres


# TODO: Função para pegar os dados de preços horarios oficiais e DESSEM e calcula a média semanal
today = date.today()
piso =  float(os.environ.get('PLD_PISO', '55.7'))
teto_diario = float(os.environ.get('PLD_TETO_DIARIO', '640.5')) 
teto_horario =  float(os.environ.get('PLD_TETO_HORARIO', '1314.02')) 


def next_friday(ref_date: date) -> date:
    output = ref_date + delta(weekday=FR)
    return output


def operative_week_month(ref_date: date) -> int:
    d = ref_date + delta(weekday=FR)
    dt_start = d.replace(day=1)
    dt_end = date(d.year, d.month, monthrange(d.year, d.month)[1])
    fridays = [d.date() for d in rrule(freq=DAILY, dtstart=dt_start, until=dt_end) if d.weekday() == 4]
    output = fridays.index(d) + 1
    return output


def get_pld_semanal(today: date = date.today()) -> pd.DataFrame:

    db_pld = Postgres(database='SeriesTemporaisCcee')

    ref_date = next_friday(ref_date=today)
    init_date = next_friday(ref_date.replace(day=1)) - delta(days=6)
    query_oficial = f"""select data_hora, submercado, pld 
                    from series_horarias.preco_oficial
                    where data_hora >= '{init_date}' and data_hora < '{today+delta(days=1)}'
                    order by data_hora asc"""
    aux = db_pld.read(query=query_oficial, to_dict=True)
    df_oficial = pd.DataFrame(data=aux)
    df_oficial['data'] = df_oficial['data_hora'].apply(lambda x: x.date())

    # Calcula PLD diário oficial
    df_oficial_dia = df_oficial.groupby(['data', 'submercado']).agg('mean').reset_index()

    previsto_em = today - delta(days=2)
    init_date = today #+ delta(days=1)
    if today.weekday() != 4:

        end_date = next_friday(init_date) #+ delta(days=1)
        query_dessem = f"""select previsto_para as data_hora, submercado, duracao, cmo
                            from series_horarias.preco_dessem 
                            where previsto_em = '{previsto_em}' and
                            previsto_para >= '{init_date}' and previsto_para < '{end_date}'
                            order by previsto_para asc"""
        aux = db_pld.read(query=query_dessem, to_dict=True)
        df_dessem = pd.DataFrame(data=aux)
        #print(df_dessem)
        df_dessem['data'] = df_dessem['data_hora'].apply(lambda x: x.date())

        # Ajusta limites horários para PLD
        df_dessem['pld'] = df_dessem['cmo'].clip(lower=piso, upper=teto_horario)

        # Calcula PLD diário DESSEM
        df_dessem_dia = df_dessem.groupby(['data', 'submercado']).apply(lambda x: np.average(x['pld'], weights=x['duracao'])).reset_index()
        df_dessem_dia.columns = ['data', 'submercado', 'pld']
        df_dessem_dia['pld'] = df_dessem_dia['pld'].clip(lower=piso, upper=teto_diario)

    else:
        df_dessem_dia = pd.DataFrame(data=[])

    # Concatena pld historico e dessem diarios
    df = pd.concat([df_oficial_dia, df_dessem_dia], axis=0)

    # Calcula PLD Médio Semanal
    df['SemOp'] = df['data'].apply(lambda x: f"DC{next_friday(x).strftime('%Y%m')}-sem{operative_week_month(x)}")
    pld_semanal = df.groupby(['SemOp', 'submercado']).agg('mean').reset_index()

    return pld_semanal
