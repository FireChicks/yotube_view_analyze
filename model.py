import pandas as pd
import os

dir = './xlsk'
file_paths = [os.path.join(dir, fileName) for fileName in os.listdir(dir)]

last_prob_list = []
forth_prob_list = []

for file_path in file_paths:
    df = pd.read_excel(file_path)
    if len(df['조회수']) >= 4:  # 최소 4개 이상의 조회수 데이터가 있는 경우에만 비교
        df = df.sort_values(by='조회수', ascending=False)
        last_prob_list.append((df['조회수'].iloc[-1] / df['조회수'].iloc[0]))
        forth_prob_list.append((df['조회수'].iloc[3] / df['조회수'].iloc[0]))

last_prob = sum(last_prob_list) / len(last_prob_list) if last_prob_list else 0.0
forth_prob = sum(forth_prob_list) / len(forth_prob_list) if forth_prob_list else 0.0

print(f"끝까지 들을 확률 : {last_prob:.2%}")
print(f"4번째 강의까지 들을 확률 : {forth_prob:.2%}")
