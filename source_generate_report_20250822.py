print('=== generate_report_v250822 ===')
print('========= preparing... ========')

import pandas as pd; print('imported pandas')
import numpy as np; print('imported numpy')
import plotly.express as px
import plotly.graph_objs as go
import plotly.subplots as sp
print('imported plotly')

# 파일 가져오기
filename = "오버워치 승패 기록표.xlsx"
original_df = pd.read_excel(filename)
print(f'loaded {filename}')

print('===== completed preparing =====')

# config
# fig() 함수에 반영
today = input('yyyymmdd: ')
today = f"{today[:4]}.{today[4:6]}.{today[6:]}."

#configured
rank_wlsum_prog = f'경쟁전 승패합계 변동 추이 ({today})'
rank_wrate_prog = f'경쟁전 승률 변동 추이 ({today})'
rank_wrate = f'경쟁전 승률 ({today})'
rank_wlsum_recent100_prog = f'경쟁전 최근 100판 승패합계 변동 추이 ({today})'
rank_wrate_recent100_prog = f'경쟁전 최근 100판 승률 변동 추이 ({today})'
rank_recent100_wrate = f'경쟁전 최근 100판 승률 ({today})'

casual_wlsum_prog = f'일반전 승패합계 변동 추이 ({today})'
casual_wrate_prog = f'일반전 승률 변동 추이 ({today})'
casual_wrate = f'일반전 승률 ({today})'
casual_wlsum_recent100_prog = f'일반전 최근 100판 승패합계 변동 추이 ({today})'
casual_wrate_recent100_prog = f'일반전 최근 100판 승률 변동 추이 ({today})'
casual_recent100_wrate = f'일반전 최근 100판 승률 ({today})'

rank_casual_wlsum_prog = f'경쟁전, 일반전 승패합계 변동 추이 ({today})'
rank_casual_wrate_prog = f'경쟁전, 일반전 승률 변동 추이 ({today})'
rank_casual_wrate = f'경쟁전, 일반전 승률 ({today})'
rank_casual_wlsum_recent100_prog = f'경쟁전, 일반전 최근 100판 승패합계 변동 추이 ({today})'
rank_casual_wrate_recent100_prog = f'경쟁전, 일반전 최근 100판 승률 변동 추이 ({today})'
rank_casual_recent100_wrate = f'경쟁전, 일반전 최근 100판 승률 ({today})'

df = original_df.copy()
print('=========generating...=========')

# 전처리: 일자 행 'yyyy-mm-dd 00:00:00' 형식으로 돼 있음. 00:00:00 제거
# 날짜 입력 실수 잡아내기
try: 
    df['일자'] = df['일자'].dt.date
except ValueError as e:
    print(e)

# 승: +1 패: -1 무: 0인 새 컬럼 생성
wlvalue = []
cnt=0
for i in df['승패'] :
    if i == '패' or i == '탈' :
        wlvalue.append(-1)
    elif i == '승' :
        wlvalue.append(1)
    elif i == '무':
        wlvalue.append(0)
    else :
        print(f'{cnt+1} Index 행 승패 열에 오타 있음')
        break
    cnt += 1

df['wlvalue'] = wlvalue

###########################################################
# 전처리함수
# 매개변수 
#   df: 데이터
#   gamemode: rank/casual/rank_casual (일반/경쟁/경쟁+일반)
# 승패값(wlvalue) 컬럼 생성(승: +1, 패: -1, 무: 0)

def pre(df, gamemode) :

    if gamemode == 'rank' :
        df_gm = df[df['게임유형']=='경쟁']
    elif gamemode == 'casual' :
        df_gm = df[df['게임유형']=='빠대']
    elif gamemode == 'rank_casual' :
        df_gm = df[(df['게임유형']=='빠대')|(df['게임유형']=='경쟁')]
        #df_gm = df_gm.reset_index().drop('index', axis=1)
    df_gm = df_gm.reset_index().drop('index', axis=1)
    
    # 승패합계(wlvalue) 컬럼 생성
    sum_wlvalue = []
    idx = 0
    for i in df_gm['wlvalue'] :
        sum_wlvalue.append(idx + i)
        idx = idx + i
    df_gm['sum_wlvalue'] = sum_wlvalue
    
    # 승률(winrate) 컬럼 생성
    winrate = []
    idx = 0
    win = 0
    los = 0
    tie = 0
    
    for i in range(0, len(df_gm)) :
        if df_gm['wlvalue'][i] == 0 :
            winrate.append(idx)
            tie += 1
        else :
            win = (df_gm['sum_wlvalue'][i]+i+1-tie)/2
            los = i+1-win-tie
            idx = round(win/(win+los+tie)*100, 2)
            winrate.append(idx)
    df_gm['winrate'] = winrate

    return df_gm

# 최근 ?판 승률(winrate_recent100) 데이터 생성 함수
def recent(df, gamemode, num=100) :
    df_gm = pre(df, gamemode)
    df_gm_recent = df_gm.tail(num).reset_index()

    # 승패합계(wlvalue) 컬럼 생성
    sum_wlvalue = []
    idx = 0
    for i in df_gm_recent['wlvalue'] :
        sum_wlvalue.append(idx + i)
        idx = idx + i
    df_gm_recent['sum_wlvalue'] = sum_wlvalue
    
    # 승률(winrate) 컬럼 생성
    winrate = []
    idx = 0
    win = 0
    los = 0
    tie = 0
    
    for i in range(0, len(df_gm_recent)) :
        if df_gm_recent['wlvalue'][i] == 0 :
            winrate.append(idx)
            tie += 1
        else :
            win = (df_gm_recent['sum_wlvalue'][i]+i+1-tie)/2
            los = i+1-win-tie
            idx = round(win/(win+los+tie)*100, 2)
            winrate.append(idx)
    df_gm_recent['winrate'] = winrate

    return df_gm_recent
    
    
###################################################################
# 그래프 생성 함수
# 매개변수
#   df: 데이터
#   gamemode: rank/casual/rank_casual (일반/경쟁/일반+경쟁)
#   index: wlsum_prog/wrate_prog/wrate (승패합계추이/승률추이/승률)

def figure(df, gamemode, index) :
    # 그래프 생성 (전처리함수 호출)
    # 그래프 이름 생성 함수(config 종속성 있음)
    title_dic = {'rank_wlsum_prog': rank_wlsum_prog, 'rank_wrate_prog': rank_wrate_prog, 'rank_wrate': rank_wrate,
                 'casual_wlsum_prog': casual_wlsum_prog, 'casual_wrate_prog': casual_wrate_prog, 'casual_wrate': casual_wrate,
                 'rank_casual_wlsum_prog': rank_casual_wlsum_prog, 'rank_casual_wrate_prog': rank_casual_wrate_prog, 'rank_casual_wrate': rank_casual_wrate}
    title = title_dic[gamemode + '_' + index]

    index_dic ={'wlsum_prog': 'sum_wlvalue', 'wrate_prog': 'winrate'}
    
    # 그래프 그리기
    df_gm = pre(df, gamemode)
    print(f'{gamemode} {index} data ready')
    if index != 'wrate' :
        fig = px.line(data_frame= df_gm, x=df_gm.index, y=index_dic[index], markers=True, width = 1000, height = 400,  text='일자')
        fig.update_traces(mode='lines')
        fig.update_layout(title = title, title_x = 0.5, hovermode = 'x unified')
        fig.update_xaxes(showgrid=True, minor_showgrid=True, ticks='inside')
        fig.update_yaxes(showgrid=True, minor_showgrid=True, ticks='inside')
        
        if index == 'wlsum_prog' :
            fig.add_hline(y = 0, line_color = 'grey')
        elif index =='wrate_prog' :
            fig.add_hline(y = 50, line_color = 'grey')
        
        return fig
    # return() 썼을 때 문자열 합성 안 됨. 해결방법 찾아야 함
    
    elif index == 'wrate' :
        print('', title, ':', len(df_gm[df_gm['승패']=='승'])/len(df_gm)*100, '%\n')

        
# 최근 (100)개 그래프 그리기 함수
def figure_recent(df, gamemode, index, num=100) : # 개선 필요. 지금은 반드시 최근 '100'개만 그래프로 만들 수 있음.
    title_dic = {'rank_wrate_prog': rank_wrate_recent100_prog, 'rank_wlsum_prog': rank_wlsum_recent100_prog, 'rank_wrate': rank_recent100_wrate,
                 'casual_wrate_prog': casual_wrate_recent100_prog, 'casual_wlsum_prog': rank_wlsum_recent100_prog, 'casual_wrate': casual_recent100_wrate,
                 'rank_casual_wrate_prog': rank_casual_wrate_recent100_prog, 'rank_casual_wlsum_prog': rank_casual_wlsum_recent100_prog, 'rank_casual_wrate': rank_casual_recent100_wrate,}
    title = title_dic[gamemode + '_' + index]

    index_dic = {'wlsum_prog': 'sum_wlvalue', 'wrate_prog': 'winrate'}

    # 그래프 그리기
    df_gm_recent = recent(df, gamemode, num)
    print(f'{gamemode} recent{num} {index} data ready')
    
    if index != 'wrate' :
        fig = px.line(data_frame=df_gm_recent, x=df_gm_recent.index, y=index_dic[index], markers=True, width = 1000, height = 400,  text='일자')
        fig.update_traces(mode='markers+lines')
        fig.update_layout(title = title, title_x = 0.5, hovermode = 'x unified')
        fig.update_xaxes(showgrid=True, minor_showgrid=True, ticks='inside')
        fig.update_yaxes(showgrid=True, minor_showgrid=True, ticks='inside')
    
        if index == 'wlsum_prog' :
            fig.add_hline(y = 0, line_color = 'grey')
        elif index =='wrate_prog' :
            fig.add_hline(y = 50, line_color = 'grey')
        
        return fig

    elif index == 'wrate' :
        print('', title, ':', len(df_gm_recent[df_gm_recent['승패']=='승'])/len(df_gm_recent)*100, '%\n')

##############################################################################################################
#############################################그래프 생성 및 저장##############################################
##############################################################################################################
print('generating figures...')
fig = sp.make_subplots(rows=6, cols=2, subplot_titles=(rank_wlsum_prog, rank_wrate_prog, rank_wlsum_recent100_prog, rank_wrate_recent100_prog,
                                                       casual_wlsum_prog, casual_wrate_prog, casual_wlsum_recent100_prog, casual_wrate_recent100_prog,
                                                       rank_casual_wlsum_prog, rank_casual_wrate_prog, rank_casual_wlsum_recent100_prog, rank_casual_wrate_recent100_prog))
# 경쟁
fig_rank_wlsum_prog = figure(df, 'rank', 'wlsum_prog')
fig_rank_wrate_prog = figure(df, 'rank', 'wrate_prog')
fig_rank_wrate = figure(df, 'rank', 'wrate')

fig_rank_recent100_wlsum_prog = figure_recent(df, 'rank', 'wlsum_prog')
fig_rank_recent100_wrate_prog = figure_recent(df, 'rank', 'wrate_prog')
fig_rank_recent100_wrate = figure_recent(df, 'rank', 'wrate')

# 빠대
fig_casual_wlsum_prog = figure(df, 'casual', 'wlsum_prog')
fig_casual_wrate_prog = figure(df, 'casual', 'wrate_prog')
fig_casual_wrate = figure(df, 'casual', 'wrate')

fig_casual_recent100_wlsum_prog = figure_recent(df, 'casual', 'wlsum_prog')
fig_casual_recent100_wrate_prog = figure_recent(df, 'casual', 'wrate_prog')
fig_casual_recent100_wrate = figure_recent(df, 'casual', 'wrate')

# 경쟁+빠대
fig_rank_casual_wlsum_prog = figure(df, 'rank_casual', 'wlsum_prog')
fig_rank_casual_wrate_prog = figure(df, 'rank_casual', 'wrate_prog')
fig_rank_casual_wrate = figure(df, 'rank_casual', 'wrate')

fig_rank_casual_recent100_wlsum_prog = figure_recent(df, 'rank_casual', 'wlsum_prog')
fig_rank_casual_recent100_wrate_prog = figure_recent(df, 'rank_casual', 'wrate_prog')
fig_rank_casual_recent100_wrate = figure_recent(df, 'rank_casual', 'wrate')

# 경쟁
for trace in fig_rank_wlsum_prog.data :
    fig.add_trace(trace, row=1, col=1)
for trace in fig_rank_wrate_prog.data :
    fig.add_trace(trace, row=1, col=2)
    
for trace in fig_rank_recent100_wlsum_prog.data :
    fig.add_trace(trace, row=2, col=1)
for trace in fig_rank_recent100_wrate_prog.data :
    fig.add_trace(trace, row=2, col=2)

# 빠대
for trace in fig_casual_wlsum_prog.data :
    fig.add_trace(trace, row=3, col=1)
for trace in fig_casual_wrate_prog.data :
    fig.add_trace(trace, row=3, col=2)
    
for trace in fig_casual_recent100_wlsum_prog.data :
    fig.add_trace(trace, row=4, col=1)
for trace in fig_casual_recent100_wrate_prog.data :
    fig.add_trace(trace, row=4, col=2)

# 경쟁+빠대
for trace in fig_rank_casual_wlsum_prog.data :
    fig.add_trace(trace, row=5, col=1)
for trace in fig_rank_casual_wrate_prog.data :
    fig.add_trace(trace, row=5, col=2)

for trace in fig_rank_casual_recent100_wlsum_prog.data :
    fig.add_trace(trace, row=6, col=1)
for trace in fig_rank_casual_recent100_wrate_prog.data :
    fig.add_trace(trace, row=6, col=2)

# 호버모드, 수평선 설정 다시 해 줘야 함
fig.update_layout(width = 1600, height = 2600, hovermode = 'x unified')
fig.add_hline(y = 0, line_color = 'grey', col = 1)
fig.add_hline(y = 50, line_color = 'grey', col = 2)

print('figures ready')
print('=====completed generation======')

fig.show()

# 저장
fig.update_layout(width = 1600, height = 2600)
fig.write_html('오버워치 승패 분석 %s.html' %today)
