import pandas as pd

# CSV dosyasını okuma
data = pd.read_csv('data.csv')

# Vardiyaları aynı olan işçileri gruplama
grouped = data.groupby('Vardiya')

# Excel dosyası oluşturma
writer = pd.ExcelWriter('sonuclar.xlsx', engine='xlsxwriter')

# Her bir grup için takım oluşturma
for name, group in grouped:
    # Yetkinlik değerlerinin toplamını hesaplama
    group['Yetkinlik Toplam'] = group['Yetkinlik 1'] + group['Yetkinlik 2'] + group['Yetkinlik 3'] + group['Yetkinlik 4']
    
    # Yetkinlik değerlerinin toplamına göre sıralama
    sorted_group = group.sort_values(by='Yetkinlik Toplam')
    
    # Takım sayısı
    team_count = 4
    
    # Takım boyutu
    team_size = len(sorted_group) // team_count
    
    # Takımları oluşturma
    teams = [[] for _ in range(team_count)]
    for i in range(team_size):
        for j in range(team_count):
            # En yüksek ve en düşük yetkinlik değerine sahip işçileri seçme
            if not sorted_group.empty:
                top_worker = sorted_group.iloc[-1]
                teams[j].append(top_worker)
                sorted_group = sorted_group.iloc[:-1]
            
            if not sorted_group.empty:
                bottom_worker = sorted_group.iloc[0]
                teams[j].append(bottom_worker)
                sorted_group = sorted_group.iloc[1:]
    
    # Takımları Excel dosyasına kaydetme
    for i, team in enumerate(teams):
        sheet_name = f'Vardiya {name} Takım {i+1}'
        team_df = pd.DataFrame(team)
        team_df.to_excel(writer, sheet_name=sheet_name, index=False)

# Excel dosyasını kaydetme
writer.save()
