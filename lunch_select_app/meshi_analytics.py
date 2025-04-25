import six
import matplotlib.pyplot as plt
import pandas as pd
from matplotlib import font_manager

font_path = "C:\\Windows\\Fonts\\msgothic.ttc"
prop = font_manager.FontProperties(fname=font_path)

df = pd.read_excel('visit_count.xlsx', sheet_name="2025年04月")

most_visited_store = df.loc[df['訪問回数'].idxmax()]
print(f"一番多く行ったお店は{most_visited_store['店舗']}で{most_visited_store['訪問回数']}回だね！")

colors = ['skyblue' if store != most_visited_store['店舗'] else 'orange' for store in df['店舗']]

plt.bar(df['店舗'], df['訪問回数'], color=colors)

for index, value in enumerate(df['訪問回数']):
    if df.iloc[index]['店舗'] == most_visited_store['店舗']:
        plt.text(index, value + 0.1, f"最多訪問！", ha='center', fontproperties=prop)

plt.xlabel('店舗', fontproperties=prop)
plt.ylabel('\n'.join("訪問回数"), rotation=0, labelpad=40, fontproperties=prop)
plt.title('今月、最もチョイスされたのはこの店だ！', fontproperties=prop)
plt.xticks(rotation=45, ha='right', fontproperties=prop)
plt.tight_layout()
plt.show()
